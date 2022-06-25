Imports System.Collections.Generic
Imports System.Globalization

Namespace Microsoft.Nautilus.Core.Undo
    Friend Class LinkedUndoTransactionImpl
        Inherits LinkedUndoTransaction

        Private linkManager As LinkedUndoTransactionManager

        Public Overrides Property Description As String

        Public Overrides Property CompletionTime As DateTime

        Public Overrides ReadOnly Property State As UndoTransactionState

        Public Overrides ReadOnly Property CanRedo As Boolean
            Get
                For Each transaction As UndoTransactionImpl In Transactions
                    If Not transaction.CanRedoInIsolation Then
                        Return False
                    End If
                Next

                Return True
            End Get
        End Property

        Public Overrides ReadOnly Property CanUndo As Boolean
            Get
                For Each transaction As UndoTransactionImpl In Transactions
                    If Not transaction.CanUndoInIsolation Then
                        Return False
                    End If
                Next

                Return True
            End Get
        End Property

        Private _transactions As List(Of UndoTransactionImpl)

        Public ReadOnly Property Transactions As IEnumerable(Of UndoTransactionImpl)
            Get
                Return _transactions
            End Get
        End Property

        Public Overrides Property Policy As ILinkedUndoTransactionPolicy

        Public Sub New(linkManager As LinkedUndoTransactionManager, description As String)
            If String.IsNullOrEmpty(description) Then
                Throw New ArgumentNullException("description", String.Format(CultureInfo.CurrentUICulture, "'{0}' called with a null value for argument '{1}'.  This value is not an allowable value.", New Object(1) {"LinkedUndoTransactionImpl", "description"}))
            End If

            If linkManager Is Nothing Then
                Throw New ArgumentNullException("linkManager", String.Format(CultureInfo.CurrentUICulture, "'{0}' called with a null value for argument '{1}'.  This value is not an allowable value.", New Object(1) {"LinkedUndoTransactionImpl", "linkManager"}))
            End If

            Me.Description = description
            CompletionTime = DateTime.MaxValue
            State = UndoTransactionState.Open
            _transactions = New List(Of UndoTransactionImpl)
            Me.linkManager = linkManager
            Policy = New NullLinkedUndoTransactionPolicy
        End Sub

        Public Overrides Sub Complete()
            If State <> 0 Then
                Throw New InvalidOperationException("Complete can only be invoked on UndoTransactions that are currently in an Open state.")
            End If

            CompletionTime = DateTime.Now
            _State = UndoTransactionState.Completed
        End Sub

        Public Overrides Sub Cancel()
            If State <> 0 Then
                Throw New InvalidOperationException("Cancel can only be invoked on UndoTransactions that are currently in an Open state.")
            End If

            _transactions.Clear()
            _State = UndoTransactionState.Canceled
        End Sub

        Public Overrides Sub Rollback()
            If State <> 0 Then
                Throw New InvalidOperationException("Rollback can only be invoked on UndoTransactions that are currently in an Open state.")
            End If

            _transactions.Clear()
            _State = UndoTransactionState.Open
        End Sub

        Private Function ComputeAllLinkedTransactionsForProcessing(getStack As Converter(Of UndoTransactionImpl, Stack(Of UndoTransaction))) As List(Of UndoTransactionImpl)
            Dim queue1 As New Queue(Of UndoTransactionImpl)(Transactions)
            Dim dictionary1 As New Dictionary(Of UndoTransactionImpl, Boolean)

            While queue1.Count > 0
                Dim undoTransactionImpl1 As UndoTransactionImpl = queue1.Dequeue()
                If dictionary1.ContainsKey(undoTransactionImpl1) OrElse undoTransactionImpl1.IsInvalid Then
                    Continue While
                End If

                dictionary1(undoTransactionImpl1) = True
                Dim array As UndoTransaction() = getStack(undoTransactionImpl1).ToArray()
                For i As Integer = 0 To array.Length - 1
                    Dim undoTransactionImpl2 As UndoTransactionImpl = TryCast(array(i), UndoTransactionImpl)
                    If undoTransactionImpl2 Is undoTransactionImpl1 Then
                        Exit For
                    End If

                    Dim linkedUndoTransaction1 As LinkedUndoTransactionImpl = linkManager.GetLinkedUndoTransaction(undoTransactionImpl2)
                    If linkedUndoTransaction1 Is Nothing Then
                        Continue For
                    End If

                    For Each transaction As UndoTransactionImpl In linkedUndoTransaction1.Transactions
                        If Not dictionary1.ContainsKey(transaction) Then
                            queue1.Enqueue(transaction)
                        End If
                    Next
                Next
            End While
            Return New List(Of UndoTransactionImpl)(dictionary1.Keys)
        End Function

        Public Overrides Sub [Do]()
            If Not CanRedo Then
                Throw New InvalidOperationException("UndoTransaction does not currently support Do, because CanRedo is false.")
            End If

            _State = UndoTransactionState.Redoing
            Dim flag As Boolean = True
            For Each transaction As UndoTransactionImpl In Transactions
                If transaction.History.State <> UndoHistoryState.Redoing AndAlso Not transaction.IsInvalid AndAlso Not Object.ReferenceEquals(transaction.History.RedoStack.Peek(), transaction) Then
                    flag = False
                    Exit For
                End If
            Next

            If flag Then
                For Each transaction2 As UndoTransactionImpl In Transactions
                    Dim undoHistoryImpl1 = TryCast(transaction2.History, UndoHistoryImpl)
                    undoHistoryImpl1.RedoInIsolation(transaction2)
                Next

            ElseIf Policy.UndoRedoWhenLinksAreNotTopmostInHistory Then
                Dim list1 As List(Of UndoTransactionImpl) = ComputeAllLinkedTransactionsForProcessing(Function(t As UndoTransactionImpl) t.History.RedoStack)
                For Each item As UndoTransactionImpl In list1
                    Dim undoHistoryImpl2 As UndoHistoryImpl = TryCast(item.History, UndoHistoryImpl)
                    undoHistoryImpl2.RedoInIsolation(item)
                Next
            End If
            _State = UndoTransactionState.Completed
        End Sub

        Public Overrides Sub Undo()
            If Not CanUndo Then
                Throw New InvalidOperationException("RedoTransaction does not currently support Undo, because CanUndo is false.")
            End If

            _State = UndoTransactionState.Undoing
            Dim flag As Boolean = True
            For Each transaction As UndoTransactionImpl In Transactions
                If transaction.History.State <> UndoHistoryState.Undoing AndAlso Not transaction.IsInvalid AndAlso Not Object.ReferenceEquals(transaction.History.UndoStack.Peek(), transaction) Then
                    flag = False
                    Exit For
                End If
            Next

            If flag Then
                For Each transaction2 As UndoTransactionImpl In Transactions
                    Dim undoHistoryImpl1 As UndoHistoryImpl = TryCast(transaction2.History, UndoHistoryImpl)
                    undoHistoryImpl1.UndoInIsolation(transaction2)
                Next

            ElseIf Policy.UndoRedoWhenLinksAreNotTopmostInHistory Then
                Dim list1 As List(Of UndoTransactionImpl) = ComputeAllLinkedTransactionsForProcessing(Function(t As UndoTransactionImpl) t.History.UndoStack)
                For Each item As UndoTransactionImpl In list1
                    Dim undoHistoryImpl2 As UndoHistoryImpl = TryCast(item.History, UndoHistoryImpl)
                    undoHistoryImpl2.UndoInIsolation(item)
                Next
            End If

            _State = UndoTransactionState.Undone
        End Sub

        Public Sub AddTransaction(transaction As UndoTransactionImpl)
            _transactions.Add(transaction)
        End Sub

        Protected Overrides Sub Dispose(disposing As Boolean)
            If disposing Then
                Select Case State
                    Case UndoTransactionState.Open
                        Cancel()
                    Case UndoTransactionState.Redoing, UndoTransactionState.Undoing, UndoTransactionState.Undone
                        Throw New InvalidOperationException("Following the Creation of an UndoTransaction but prior to calling Dispose, a Do() or Undo() was called illegally, and now the transaction cannot be closed properly.")
                End Select
            End If

            linkManager.EndLinkedUndoTransaction(Me)
        End Sub
    End Class
End Namespace
