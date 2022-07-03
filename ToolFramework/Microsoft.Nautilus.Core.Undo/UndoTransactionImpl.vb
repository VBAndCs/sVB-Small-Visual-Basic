Imports System.Collections.Generic
Imports System.Globalization

Namespace Microsoft.Nautilus.Core.Undo
    Friend Class UndoTransactionImpl
        Inherits UndoTransaction

        Private primitives As List(Of IUndoPrimitive)

        Friend ReadOnly Property IsInvalid As Boolean
            Get
                Return _State = UndoTransactionState.Invalid
            End Get
        End Property

        Public Overrides Property Description As String

        Public Overrides Property CompletionTime As DateTime

        Public Overrides Property IsHidden As Boolean

        Public Overrides ReadOnly Property State As UndoTransactionState

        Private ReadOnly _history As UndoHistoryImpl

        Public Overrides ReadOnly Property History As UndoHistory
            Get
                Return _history
            End Get
        End Property

        Public Overrides ReadOnly Property UndoPrimitives As IList(Of IUndoPrimitive)
            Get
                Return primitives
            End Get
        End Property

        Dim _parent As UndoTransactionImpl

        Public Overrides ReadOnly Property Parent As UndoTransaction
            Get
                Return _parent
            End Get
        End Property

        Public Overrides ReadOnly Property CanRedo As Boolean
            Get
                Dim undoHistoryRegistry1 = _history.UndoHistoryRegistry
                If undoHistoryRegistry1.LinkedUndoTransactionManager.Linked(Me) Then
                    Return undoHistoryRegistry1.LinkedUndoTransactionManager.CanRedoTransaction(Me)
                End If

                Return CanRedoInIsolation
            End Get
        End Property

        Public ReadOnly Property CanRedoInIsolation As Boolean
            Get
                If _State = UndoTransactionState.Invalid Then
                    Return True
                End If

                If _State <> UndoTransactionState.Undone Then
                    Return False
                End If

                For Each undoPrimitive As IUndoPrimitive In UndoPrimitives
                    If Not undoPrimitive.CanRedo Then
                        Return False
                    End If
                Next

                Return True
            End Get
        End Property

        Public Overrides ReadOnly Property CanUndo As Boolean
            Get
                Dim undoHistoryRegistry1 = _history.UndoHistoryRegistry
                If undoHistoryRegistry1.LinkedUndoTransactionManager.Linked(Me) Then
                    Return undoHistoryRegistry1.LinkedUndoTransactionManager.CanUndoTransaction(Me)
                End If

                Return CanUndoInIsolation
            End Get
        End Property

        Public ReadOnly Property CanUndoInIsolation As Boolean
            Get
                If _State = UndoTransactionState.Invalid Then
                    Return True
                End If

                If _State <> UndoTransactionState.Completed Then
                    Return False
                End If

                For Each undoPrimitive As IUndoPrimitive In UndoPrimitives
                    If Not undoPrimitive.CanUndo Then
                        Return False
                    End If
                Next

                Return True
            End Get
        End Property

        Dim _mergePolicy As IMergeUndoTransactionPolicy

        Public Overrides Property MergePolicy As IMergeUndoTransactionPolicy
            Get
                Return _mergePolicy
            End Get

            Set(value As IMergeUndoTransactionPolicy)
                If value Is Nothing Then
                    Throw New ArgumentNullException("value")
                End If

                _mergePolicy = value
            End Set
        End Property

        Public Sub New(history As UndoHistoryImpl, parent As UndoTransaction, description As String, isHidden As Boolean)
            If history Is Nothing Then
                Throw New ArgumentNullException("history", String.Format(CultureInfo.CurrentUICulture, "'{0}' called with a null value for argument '{1}'.  This value is not an allowable value.", New Object(1) {"UndoTransactionImpl", "history"}))
            End If

            If String.IsNullOrEmpty(description) Then
                Throw New ArgumentNullException("description", String.Format(CultureInfo.CurrentUICulture, "'{0}' called with a null value for argument '{1}'.  This value is not an allowable value.", New Object(1) {"UndoTransactionImpl", "description"}))
            End If

            _history = history
            _parent = TryCast(parent, UndoTransactionImpl)

            If Me.Parent Is Nothing AndAlso parent IsNot Nothing Then
                Throw New ArgumentException("UndoTransactionImpl must be created with a reference to a valid UndoTransactionImpl")
            End If

            _Description = description
            _IsHidden = isHidden

            _State = UndoTransactionState.Open
            _CompletionTime = DateTime.MaxValue
            primitives = New List(Of IUndoPrimitive)
            MergePolicy = NullMergeUndoTransactionPolicy.Instance
        End Sub

        Friend Sub Invalidate()
            _State = UndoTransactionState.Invalid
        End Sub

        Public Overrides Sub Complete()
            If _State <> 0 Then
                Throw New InvalidOperationException("Complete can only be invoked on UndoTransactions that are currently in an Open state.")
            End If

            _CompletionTime = DateTime.Now
            _State = UndoTransactionState.Completed
            FlattenPrimitivesToParent()
            ReportCompletionToLinkManager()
        End Sub

        Public Sub FlattenPrimitivesToParent()
            If _parent IsNot Nothing Then
                _parent.CopyPrimitivesFrom(Me)
                primitives.Clear()
            End If
        End Sub

        Public Sub ReportCompletionToLinkManager()
            If Parent Is Nothing Then
                Dim undoHistoryRegistry1 As UndoHistoryRegistryImpl = _history.UndoHistoryRegistry
                undoHistoryRegistry1.LinkedUndoTransactionManager.ReportTransactionCompletion(Me)
            End If
        End Sub

        Public Sub CopyPrimitivesFrom(transaction As UndoTransactionImpl)
            For Each undoPrimitive In transaction.UndoPrimitives
                AddUndo(undoPrimitive)
            Next
        End Sub

        Public Overrides Sub Cancel()
            If _State <> 0 Then
                Throw New InvalidOperationException("Cancel can only be invoked on UndoTransactions that are currently in an Open state.")
            End If

            For num As Integer = primitives.Count - 1 To 0 Step -1
                primitives(num).Undo()
            Next

            primitives.Clear()
            _State = UndoTransactionState.Canceled
        End Sub

        Public Overrides Sub Rollback()
            If _State <> 0 Then
                Throw New InvalidOperationException("Rollback can only be invoked on UndoTransactions that are currently in an Open state.")
            End If

            For num As Integer = primitives.Count - 1 To 0 Step -1
                primitives(num).Undo()
            Next

            primitives.Clear()
            _State = UndoTransactionState.Open
        End Sub

        Public Overrides Sub AddUndo(undo1 As IUndoPrimitive)
            If _State <> 0 Then
                Throw New InvalidOperationException("AddUndo can only be invoked on UndoTransactions that are currently in an Open state.")
            End If

            primitives.Add(undo1)
            undo1.Parent = Me
            MergeMostRecentUndoPrimitive()
        End Sub

        Protected Sub MergeMostRecentUndoPrimitive()
            If primitives.Count < 2 Then
                Return
            End If

            Dim undoPrimitive = primitives(primitives.Count - 1)
            If undoPrimitive.MergeWithPreviousOnly Then
                Dim undoPrimitive2 = primitives(primitives.Count - 2)
                If undoPrimitive.GetType() = undoPrimitive2.GetType() AndAlso undoPrimitive.CanMerge(undoPrimitive2) Then
                    Dim item = undoPrimitive.Merge(undoPrimitive2)
                    primitives.RemoveRange(primitives.Count - 2, 2)
                    primitives.Add(item)
                End If

                Return
            End If

            Dim undoPrimitive3 As IUndoPrimitive = Nothing
            Dim index As Integer = -1
            For num = primitives.Count - 2 To 0 Step -1
                If undoPrimitive.GetType() = primitives(num).GetType() AndAlso undoPrimitive.CanMerge(primitives(num)) Then
                    undoPrimitive3 = primitives(num)
                    index = num
                    Exit For
                End If
            Next

            If undoPrimitive3 IsNot Nothing Then
                Dim item2 As IUndoPrimitive = undoPrimitive.Merge(undoPrimitive3)
                primitives.RemoveRange(primitives.Count - 1, 1)
                primitives.RemoveRange(index, 1)
                primitives.Add(item2)
            End If
        End Sub

        Public Overrides Sub [Do]()
            Dim undoHistoryRegistry1 = _history.UndoHistoryRegistry
            If undoHistoryRegistry1.LinkedUndoTransactionManager.Linked(Me) Then
                undoHistoryRegistry1.LinkedUndoTransactionManager.DoTransaction(Me)
                DoInIsolation()
            Else
                DoInIsolation()
            End If
        End Sub

        Public Sub DoInIsolation()
            If _State <> UndoTransactionState.Invalid Then
                If Not CanRedoInIsolation Then
                    Throw New InvalidOperationException("UndoTransaction does not currently support Do, because CanRedo is false.")
                End If

                _State = UndoTransactionState.Redoing
                For i As Integer = 0 To primitives.Count - 1
                    primitives(i).Do()
                Next

                _State = UndoTransactionState.Completed
            End If
        End Sub

        Public Overrides Sub Undo()
            Dim undoHistoryRegistry1 = _history.UndoHistoryRegistry
            If undoHistoryRegistry1.LinkedUndoTransactionManager.Linked(Me) Then
                undoHistoryRegistry1.LinkedUndoTransactionManager.UndoTransaction(Me)
                UndoInIsolation()
            Else
                UndoInIsolation()
            End If
        End Sub

        Public Sub UndoInIsolation()
            If _State <> UndoTransactionState.Invalid Then
                If Not CanUndoInIsolation Then
                    Throw New InvalidOperationException("RedoTransaction does not currently support Undo, because CanUndo is false.")
                End If

                _State = UndoTransactionState.Undoing
                For num As Integer = primitives.Count - 1 To 0 Step -1
                    primitives(num).Undo()
                Next

                _State = UndoTransactionState.Undone
            End If
        End Sub

        Public Overrides Function GetEstimatedSize() As Long
            Dim num As Long = 0L
            For Each primitive In primitives
                num += primitive.GetEstimatedSize()
            Next
            Return num
        End Function

        Protected Overrides Sub Dispose(disposing As Boolean)
            If disposing Then
                Select Case _State
                    Case UndoTransactionState.Open
                        Cancel()
                    Case UndoTransactionState.Redoing, UndoTransactionState.Undoing, UndoTransactionState.Undone
                        Throw New InvalidOperationException("Following the Creation of an UndoTransaction but prior to calling Dispose, a Do() or Undo() was called illegally, and now the transaction cannot be closed properly.")
                End Select
            End If
            _history.EndTransaction(Me)
        End Sub
    End Class
End Namespace
