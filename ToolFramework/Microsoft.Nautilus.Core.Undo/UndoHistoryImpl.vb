Imports System.Collections.Generic
Imports System.Globalization
Imports Microsoft.Nautilus.Core.Undo.DataExports
Imports System.Runtime.InteropServices

Namespace Microsoft.Nautilus.Core.Undo
    Friend Class UndoHistoryImpl
        Inherits UndoHistory

        Private _currentTransaction As UndoTransactionImpl
        Private markers As Dictionary(Of UndoTransactionMarker, UndoTransactionMarkerReferenceValuePair)
        Private activeUndoOperationPrimitive As DelegatedUndoPrimitiveImpl

        Friend ReadOnly Property UndoHistoryRegistry As UndoHistoryRegistryImpl

        Public Overrides ReadOnly Property UndoStack As Stack(Of UndoTransaction)

        Public Overrides ReadOnly Property RedoStack As Stack(Of UndoTransaction)

        Public Overrides ReadOnly Property CanUndo As Boolean
            Get
                If _UndoStack.Count > 0 Then
                    Return _UndoStack.Peek().CanUndo
                End If

                Return False
            End Get
        End Property

        Public Overrides ReadOnly Property CanRedo As Boolean
            Get
                If _RedoStack.Count > 0 Then
                    Return _RedoStack.Peek().CanRedo
                End If

                Return False
            End Get
        End Property

        Public Overrides ReadOnly Property UndoDescription As String
            Get
                If _UndoStack.Count > 0 Then
                    Return _UndoStack.Peek().Description
                End If

                Return "Can't Undo"
            End Get
        End Property

        Public Overrides ReadOnly Property RedoDescription As String
            Get
                If _RedoStack.Count > 0 Then
                    Return _RedoStack.Peek().Description
                End If

                Return "Can't Redo"
            End Get
        End Property

        Public Overrides ReadOnly Property CurrentTransaction As UndoTransaction
            Get
                Return _currentTransaction
            End Get
        End Property

        Public Overrides ReadOnly Property State As UndoHistoryState

        Public Sub New(undoHistoryRegistry As UndoHistoryRegistryImpl)
            _UndoHistoryRegistry = undoHistoryRegistry
            _currentTransaction = Nothing
            _UndoStack = New Stack(Of UndoTransaction)
            _RedoStack = New Stack(Of UndoTransaction)
            markers = New Dictionary(Of UndoTransactionMarker, UndoTransactionMarkerReferenceValuePair)
            activeUndoOperationPrimitive = Nothing
            _State = UndoHistoryState.Idle
        End Sub

        Public Overrides Function GetEstimatedSize() As Long
            Dim num As Long = 0L
            For Each item As UndoTransaction In _UndoStack
                num += item.GetEstimatedSize()
            Next

            For Each item2 As UndoTransaction In _RedoStack
                num += item2.GetEstimatedSize()
            Next
            Return num
        End Function

        Public Overrides Function CreateTransaction(description As String) As UndoTransaction
            If String.IsNullOrEmpty(description) Then
                Throw New ArgumentNullException("description", String.Format(CultureInfo.CurrentUICulture, "'{0}' called with a null value for argument '{1}'.  This value is not an allowable value.", New Object(1) {"CreateTransaction", "description"}))
            End If

            Return CreateTransaction(description, CurrentTransaction IsNot Nothing AndAlso CurrentTransaction.IsHidden)
        End Function

        Public Overrides Function CreateTransaction(description As String, isHidden As Boolean) As UndoTransaction
            If String.IsNullOrEmpty(description) Then
                Throw New ArgumentNullException("description", String.Format(CultureInfo.CurrentUICulture, "'{0}' called with a null value for argument '{1}'.  This value is not an allowable value.", New Object(1) {"CreateTransaction", "description"}))
            End If

            If CurrentTransaction IsNot Nothing AndAlso CurrentTransaction.State <> 0 Then
                Throw New InvalidOperationException("The UndoHistory cannot fulfill your request to create a new transaction, because the current transaction has been Completed but not yet been Disposed. You should Complete() an UndoTransaction as the last step in the CreateTransaction using block.")
            End If

            If Not isHidden Then
                If CurrentTransaction Is Nothing Then
                    For Each item As UndoTransactionImpl In _RedoStack
                        _UndoHistoryRegistry.LinkedUndoTransactionManager.BreakLinkToUndoTransaction(item)
                        item.Invalidate()
                    Next
                    _RedoStack.Clear()

                ElseIf CurrentTransaction.IsHidden Then
                    Throw New InvalidOperationException("A nested UndoTransaction which is not marked IsHidden cannot be created within an UndoTransaction which is marked IsHidden.")
                End If
            End If

            _currentTransaction = New UndoTransactionImpl(Me, CurrentTransaction, description, isHidden)
            Return _currentTransaction
        End Function

        Public Overrides Sub Undo(count1 As Integer)
            If count1 <= 0 Then
                Throw New ArgumentException(String.Format(CultureInfo.CurrentUICulture, "The UndoHistory operation '{0}' only accepts positive count values, but {1} was received.", New Object(1) {"Undo", count1}), "count")
            End If

            If count1 > _UndoStack.Count Then
                Throw New InvalidOperationException(String.Format(CultureInfo.CurrentUICulture, "An UndoHistory received a request to {0} {1} transactions, which is more transactions than currently exist in the stack.", New Object(1) {"undo", count1}))
            End If

            Dim originalState As UndoHistoryState = _State
            _State = UndoHistoryState.Undoing
            Using New AutoEnclose(Sub() _State = originalState)
                While count1 > 0
                    If Not CanUndo Then
                        Throw New InvalidOperationException("The UndoHistory cannot fulfill your request to Undo because at least one of the UndoPrimitives does not permit it.")
                    End If

                    Dim undoTransaction1 = _UndoStack.Pop()
                    undoTransaction1.Undo()
                    _RedoStack.Push(undoTransaction1)
                    RaiseUndoRedoHappened(_State, undoTransaction1)
                    count1 -= 1
                End While

                Dim undoTransaction2 = (If((_UndoStack.Count > 0), _UndoStack.Peek(), Nothing))

                While undoTransaction2 IsNot Nothing AndAlso
                           undoTransaction2.IsHidden AndAlso
                           undoTransaction2.CanUndo AndAlso
                           Not _UndoHistoryRegistry.LinkedUndoTransactionManager.Linked(CType(undoTransaction2, UndoTransactionImpl))

                    Dim undoTransaction3 = _UndoStack.Pop()
                    undoTransaction3.Undo()
                    _RedoStack.Push(undoTransaction3)
                    RaiseUndoRedoHappened(_State, undoTransaction3)
                    undoTransaction2 = (If((_UndoStack.Count > 0), _UndoStack.Peek(), Nothing))
                End While
            End Using
        End Sub

        Public Sub UndoInIsolation(transaction As UndoTransactionImpl)
            Dim originalState As UndoHistoryState = _State
            _State = UndoHistoryState.Undoing
            Dim [end] As AutoEncloseDelegate = Sub() _State = originalState

            Using New AutoEnclose([end])
                If _UndoStack.Contains(transaction) Then
                    Dim undoTransactionImpl1 As UndoTransactionImpl = Nothing
                    While undoTransactionImpl1 IsNot transaction
                        Dim undoTransactionImpl2 As UndoTransactionImpl = TryCast(_UndoStack.Pop(), UndoTransactionImpl)
                        undoTransactionImpl2.UndoInIsolation()
                        _RedoStack.Push(undoTransactionImpl2)
                        RaiseUndoRedoHappened(_State, undoTransactionImpl2)
                        undoTransactionImpl1 = undoTransactionImpl2
                    End While

                    Dim undoTransaction1 As UndoTransaction = (If((_UndoStack.Count > 0), _UndoStack.Peek(), Nothing))
                    While undoTransaction1 IsNot Nothing AndAlso undoTransaction1.IsHidden AndAlso undoTransaction1.CanUndo AndAlso Not _UndoHistoryRegistry.LinkedUndoTransactionManager.Linked(CType(undoTransaction1, UndoTransactionImpl))
                        Dim undoTransactionImpl3 = TryCast(_UndoStack.Pop(), UndoTransactionImpl)
                        undoTransactionImpl3.UndoInIsolation()
                        _RedoStack.Push(undoTransactionImpl3)
                        RaiseUndoRedoHappened(_State, undoTransactionImpl3)
                        undoTransaction1 = (If((_UndoStack.Count > 0), _UndoStack.Peek(), Nothing))
                    End While
                End If
            End Using
        End Sub

        Public Overrides Sub Redo(count As Integer)
            If count <= 0 Then
                Throw New ArgumentException(String.Format(CultureInfo.CurrentUICulture, "The UndoHistory operation '{0}' only accepts positive count values, but {1} was received.", New Object(1) {"Redo", count}), "count")
            End If

            If count > _RedoStack.Count Then
                Throw New InvalidOperationException(String.Format(CultureInfo.CurrentUICulture, "An UndoHistory received a request to {0} {1} transactions, which is more transactions than currently exist in the stack.", New Object(1) {"redo", count}))
            End If

            Dim originalState As UndoHistoryState = _State
            _State = UndoHistoryState.Redoing
            Using New AutoEnclose(Sub() _State = originalState)
                While count > 0
                    If Not CanRedo Then
                        Throw New InvalidOperationException("The UndoHistory cannot fulfill your request to Redo because at least one of the UndoPrimitives does not permit it.")
                    End If

                    Dim undoTransaction1 As UndoTransaction = _RedoStack.Pop()
                    undoTransaction1.Do()
                    _UndoStack.Push(undoTransaction1)
                    RaiseUndoRedoHappened(_State, undoTransaction1)
                    count -= 1
                End While

                Dim undoTransaction2 = If((_RedoStack.Count > 0), _RedoStack.Peek(), Nothing)

                While undoTransaction2 IsNot Nothing AndAlso undoTransaction2.IsHidden AndAlso undoTransaction2.CanRedo AndAlso Not _UndoHistoryRegistry.LinkedUndoTransactionManager.Linked(CType(undoTransaction2, UndoTransactionImpl))
                    Dim undoTransaction3 = _RedoStack.Pop()
                    undoTransaction3.Do()
                    _UndoStack.Push(undoTransaction3)
                    RaiseUndoRedoHappened(_State, undoTransaction3)
                    undoTransaction2 = (If((_RedoStack.Count > 0), _RedoStack.Peek(), Nothing))
                End While

            End Using
        End Sub

        Public Sub RedoInIsolation(transaction As UndoTransactionImpl)
            Dim originalState As UndoHistoryState = _State
            _State = UndoHistoryState.Redoing
            Dim [end] As AutoEncloseDelegate = Sub() _State = originalState

            Using New AutoEnclose([end])
                If _RedoStack.Contains(transaction) Then
                    Dim undoTransactionImpl1 As UndoTransactionImpl = Nothing
                    While undoTransactionImpl1 IsNot transaction
                        Dim undoTransactionImpl2 = TryCast(_RedoStack.Pop(), UndoTransactionImpl)
                        undoTransactionImpl2.DoInIsolation()
                        _UndoStack.Push(undoTransactionImpl2)
                        RaiseUndoRedoHappened(_State, undoTransactionImpl2)
                        undoTransactionImpl1 = undoTransactionImpl2
                    End While

                    Dim undoTransaction1 = (If((_RedoStack.Count > 0), _RedoStack.Peek(), Nothing))
                    While undoTransaction1 IsNot Nothing AndAlso undoTransaction1.IsHidden AndAlso undoTransaction1.CanUndo AndAlso Not _UndoHistoryRegistry.LinkedUndoTransactionManager.Linked(CType(undoTransaction1, UndoTransactionImpl))
                        Dim undoTransactionImpl3 = TryCast(_RedoStack.Pop(), UndoTransactionImpl)
                        undoTransactionImpl3.DoInIsolation()
                        _UndoStack.Push(undoTransactionImpl3)
                        RaiseUndoRedoHappened(_State, undoTransactionImpl3)
                        undoTransaction1 = (If((_RedoStack.Count > 0), _RedoStack.Peek(), Nothing))
                    End While
                End If
            End Using
        End Sub

        Public Overrides Sub [Do](undoPrimitive As IUndoPrimitive)
            If undoPrimitive Is Nothing Then
                Throw New ArgumentNullException("undoPrimitive", String.Format(CultureInfo.CurrentUICulture, "'{0}' called with a null value for argument '{1}'.  This value is not an allowable value.", New Object(1) {"Do", "undoPrimitive"}))
            End If

            If CurrentTransaction Is Nothing OrElse CurrentTransaction.State <> 0 Then
                Throw New InvalidOperationException("There must exist an open UndoTransaction in the UndoHistory to proceed with this add.")
            End If

            undoPrimitive.Do()
            CurrentTransaction.AddUndo(undoPrimitive)
        End Sub

        Public Sub ForwardToUndoOperation(primitive As DelegatedUndoPrimitiveImpl)
            If activeUndoOperationPrimitive IsNot Nothing Then
                Throw New InvalidOperationException
            End If
            activeUndoOperationPrimitive = primitive
        End Sub

        Public Sub EndForwardToUndoOperation(primitive As DelegatedUndoPrimitiveImpl)
            If activeUndoOperationPrimitive IsNot primitive Then
                Throw New InvalidOperationException
            End If
            activeUndoOperationPrimitive = Nothing
        End Sub

        Protected Overloads Sub AddUndo(operation As UndoableOperationCurried)
            If operation Is Nothing Then
                Throw New ArgumentNullException("operation")
            End If

            If activeUndoOperationPrimitive IsNot Nothing Then
                activeUndoOperationPrimitive.AddOperation(operation)
                Return
            End If

            If CurrentTransaction IsNot Nothing AndAlso CurrentTransaction.State = UndoTransactionState.Open Then
                Dim undo1 As New DelegatedUndoPrimitiveImpl(Me, CType(CurrentTransaction, UndoTransactionImpl), operation)
                CurrentTransaction.AddUndo(undo1)
                Return
            End If

            Throw New InvalidOperationException("There must exist an open UndoTransaction in the UndoHistory to proceed with this add.")
        End Sub

        Public Overrides Sub AddUndo(operation As UndoableOperationVoid)
            AddUndo(Sub() operation(Me))
        End Sub

        Public Overrides Sub AddUndo(Of P0)(operation As UndoableOperation(Of P0), p01 As P0)
            AddUndo(Sub() operation(Me, p01))
        End Sub

        Public Overrides Sub AddUndo(Of P0, P1)(operation As UndoableOperation(Of P0, P1), p01 As P0, p11 As P1)
            AddUndo(Sub() operation(Me, p01, p11))
        End Sub

        Public Overrides Sub AddUndo(Of P0, P1, P2)(operation As UndoableOperation(Of P0, P1, P2), p01 As P0, p11 As P1, p21 As P2)
            AddUndo(Sub() operation(Me, p01, p11, p21))
        End Sub

        Public Overrides Sub AddUndo(Of P0, P1, P2, P3)(operation As UndoableOperation(Of P0, P1, P2, P3), p01 As P0, p11 As P1, p21 As P2, p31 As P3)
            AddUndo(Sub() operation(Me, p01, p11, p21, p31))
        End Sub

        Public Overrides Sub AddUndo(Of P0, P1, P2, P3, P4)(operation As UndoableOperation(Of P0, P1, P2, P3, P4), p01 As P0, p11 As P1, p21 As P2, p31 As P3, p41 As P4)
            AddUndo(Sub() operation(Me, p01, p11, p21, p31, p41))
        End Sub

        Public Overrides Sub ReplaceMarker(marker As UndoTransactionMarker, transaction As UndoTransaction, _value As Object)
            If marker Is Nothing Then
                Throw New ArgumentNullException("marker")
            End If

            If Not ValidTransactionForMarkers(transaction) Then
                Throw New InvalidOperationException
            End If

            markers(marker) = New UndoTransactionMarkerReferenceValuePair(TryCast(transaction, UndoTransactionImpl), _value)
        End Sub

        Public Overrides Function GetMarker(marker As UndoTransactionMarker, transaction As UndoTransaction) As Object
            If marker Is Nothing Then
                Throw New ArgumentNullException("marker")
            End If

            Dim undoTransaction1 As UndoTransaction = Nothing
            Dim result As Object = Nothing
            If markers.ContainsKey(marker) Then
                undoTransaction1 = markers(marker).Transaction
                result = markers(marker).Value
                If Not ValidTransactionForMarkers(undoTransaction1) Then
                    markers.Remove(marker)
                    undoTransaction1 = Nothing
                    result = Nothing
                End If
            End If
            Return result
        End Function

        Public Overrides Sub ClearMarker(marker As UndoTransactionMarker)
            If marker Is Nothing Then
                Throw New ArgumentNullException("marker")
            End If
            markers(marker) = New UndoTransactionMarkerReferenceValuePair(Nothing, Nothing)
        End Sub

        Public Overrides Function FindMarker(marker As UndoTransactionMarker) As UndoTransaction
            If marker Is Nothing Then
                Throw New ArgumentNullException("marker")
            End If

            Dim undoTransaction1 As UndoTransaction = Nothing
            If markers.ContainsKey(marker) Then
                undoTransaction1 = markers(marker).Transaction
                If Not ValidTransactionForMarkers(undoTransaction1) Then
                    markers.Remove(marker)
                    undoTransaction1 = Nothing
                End If
            End If
            Return undoTransaction1
        End Function

        Public Overrides Function TryFindMarkerOnTop(marker As UndoTransactionMarker, <Out> ByRef _value As Object) As Boolean
            If marker Is Nothing Then
                Throw New ArgumentNullException("marker")
            End If

            Dim undoTransaction1 As UndoTransaction = Nothing
            Dim obj As Object = Nothing
            If markers.ContainsKey(marker) Then
                undoTransaction1 = markers(marker).Transaction
                obj = markers(marker).Value
                If Not ValidTransactionForMarkers(undoTransaction1) Then
                    markers.Remove(marker)
                    undoTransaction1 = Nothing
                    obj = Nothing
                End If
            End If

            If (_UndoStack.Count > 0 AndAlso undoTransaction1 Is _UndoStack.Peek()) OrElse (_UndoStack.Count = 0 AndAlso undoTransaction1 Is Nothing) Then
                _value = obj
                Return True
            End If

            _value = Nothing
            Return False
        End Function

        Public Overrides Sub ReplaceMarkerOnTop(marker As UndoTransactionMarker, _value As Object)
            If marker Is Nothing Then
                Throw New ArgumentNullException("marker")
            End If

            Dim undoTransaction1 = (If((_UndoStack.Count > 0), _UndoStack.Peek(), Nothing))
            If Not ValidTransactionForMarkers(undoTransaction1) Then
                Throw New InvalidOperationException
            End If

            markers(marker) = New UndoTransactionMarkerReferenceValuePair(TryCast(undoTransaction1, UndoTransactionImpl), _value)
        End Sub

        Public Sub EndTransaction(transaction As UndoTransaction)
            If CurrentTransaction IsNot transaction Then
                _currentTransaction = Nothing
                Throw New InvalidOperationException("There has been an attempt to end the creation of a new UndoTransaction which is not the most recently created. You can fix this issue by using the IDispose/using coding pattern on UndoHistory.CreateTransaction.")
            End If

            If _currentTransaction.State = UndoTransactionState.Completed AndAlso CurrentTransaction.Parent Is Nothing Then
                MergeOrPushToUndoStack(CType(CurrentTransaction, UndoTransactionImpl))
            End If
            _currentTransaction = TryCast(CurrentTransaction.Parent, UndoTransactionImpl)
        End Sub

        Private Sub MergeOrPushToUndoStack(transaction As UndoTransactionImpl)
            If _UndoStack.Count > 0 AndAlso ProceedWithMerge(transaction, TryCast(_UndoStack.Peek(), UndoTransactionImpl)) Then
                Dim undoTransactionImpl1 As UndoTransactionImpl = TryCast(_UndoStack.Pop(), UndoTransactionImpl)
                Dim undoTransactionImpl2 As New UndoTransactionImpl(Me, Nothing, transaction.Description, transaction.IsHidden)
                undoTransactionImpl2.CopyPrimitivesFrom(undoTransactionImpl1)
                undoTransactionImpl2.CopyPrimitivesFrom(transaction)
                undoTransactionImpl2.MergePolicy = transaction.MergePolicy
                transaction.MergePolicy.CompleteTransactionMerge(transaction, undoTransactionImpl1, undoTransactionImpl2)
                undoTransactionImpl2.Complete()
                Dim list1 As New List(Of UndoTransactionMarker)
                For Each key As UndoTransactionMarker In markers.Keys
                    If Object.ReferenceEquals(markers(key).Transaction, undoTransactionImpl1) OrElse Object.ReferenceEquals(markers(key).Transaction, transaction) Then
                        list1.Add(key)
                    End If
                Next

                For Each item As UndoTransactionMarker In list1
                    markers(item) = New UndoTransactionMarkerReferenceValuePair(undoTransactionImpl2, markers(item).Value)
                Next
                _UndoStack.Push(undoTransactionImpl2)
            Else
                _UndoStack.Push(transaction)
            End If
        End Sub

        Public Function ValidTransactionForMarkers(transaction As UndoTransaction) As Boolean
            If transaction IsNot Nothing AndAlso CurrentTransaction IsNot transaction Then
                If transaction.History Is Me Then
                    Return transaction.State <> UndoTransactionState.Invalid
                End If
                Return False
            End If
            Return True
        End Function

        Private Function ProceedWithMerge(transaction1 As UndoTransactionImpl, transaction2 As UndoTransactionImpl) As Boolean
            If transaction1.MergePolicy IsNot Nothing AndAlso transaction2.MergePolicy IsNot Nothing AndAlso _UndoHistoryRegistry.LinkedUndoTransactionManager.AreTransactionsSafeToMerge(transaction1, transaction2) AndAlso transaction1.MergePolicy.TestCompatiblePolicy(transaction2.MergePolicy) AndAlso transaction1.MergePolicy.CanMerge(transaction1, transaction2) AndAlso transaction1.CanUndo Then
                Return transaction2.CanUndo
            End If
            Return False
        End Function

    End Class
End Namespace
