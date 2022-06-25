Imports System.Collections.Generic
Imports System.Globalization

Namespace Microsoft.Nautilus.Core.Undo
    Friend Class LinkedUndoTransactionManager

        Private currentLinkedUndoTransaction As LinkedUndoTransactionImpl
        Private completedLinkedUndoTransactions As New List(Of LinkedUndoTransactionImpl)
        Private transactionsLinkedTo As New Dictionary(Of UndoTransactionImpl, LinkedUndoTransactionImpl)

        Public Function CreateLinkedUndoTransaction(description As String) As LinkedUndoTransactionImpl
            If String.IsNullOrEmpty(description) Then
                Throw New ArgumentNullException("description", String.Format(CultureInfo.CurrentUICulture, "'{0}' called with a null value for argument '{1}'.  This value is not an allowable value.", New Object(1) {"CreateTransaction", "description"}))
            End If

            If currentLinkedUndoTransaction IsNot Nothing Then
                Throw New InvalidOperationException("Only one open linked transaction may exist at a time, and there is already an open LinkedUndoTransaction.")
            End If

            currentLinkedUndoTransaction = New LinkedUndoTransactionImpl(Me, description)
            Return currentLinkedUndoTransaction
        End Function

        Public Sub EndLinkedUndoTransaction(linkedTransaction As LinkedUndoTransactionImpl)
            If currentLinkedUndoTransaction IsNot linkedTransaction Then
                currentLinkedUndoTransaction = Nothing
                Throw New InvalidOperationException("There has been an attempt to end the creation of a new UndoTransaction which is not the most recently created. You can fix this issue by using the IDispose/using coding pattern on UndoHistory.CreateTransaction.")
            End If

            If currentLinkedUndoTransaction.State = UndoTransactionState.Completed Then
                For Each transaction As UndoTransactionImpl In currentLinkedUndoTransaction.Transactions
                    transactionsLinkedTo.Add(transaction, currentLinkedUndoTransaction)
                Next

                completedLinkedUndoTransactions.Add(currentLinkedUndoTransaction)
            End If

            currentLinkedUndoTransaction = Nothing
        End Sub

        Public Sub ReportTransactionCompletion(transaction As UndoTransactionImpl)
            If transaction Is Nothing Then
                Throw New ArgumentNullException("transaction", String.Format(CultureInfo.CurrentUICulture, "'{0}' called with a null value for argument '{1}'.  This value is not an allowable value.", New Object(1) {"ReportTransactionCompletion", "transaction"}))
            End If
            If currentLinkedUndoTransaction IsNot Nothing Then
                currentLinkedUndoTransaction.AddTransaction(transaction)
            End If
        End Sub

        Public Function Linked(transaction As UndoTransactionImpl) As Boolean
            Return transactionsLinkedTo.ContainsKey(transaction)
        End Function

        Public Function CanRedoTransaction(transaction As UndoTransactionImpl) As Boolean
            Dim _value As LinkedUndoTransactionImpl = Nothing
            If transactionsLinkedTo.TryGetValue(transaction, _value) Then
                Return _value.CanRedo
            End If
            Return transaction.CanRedoInIsolation
        End Function

        Public Function CanUndoTransaction(transaction As UndoTransactionImpl) As Boolean
            Dim _value As LinkedUndoTransactionImpl = Nothing
            If transactionsLinkedTo.TryGetValue(transaction, _value) Then
                Return _value.CanUndo
            End If
            Return transaction.CanUndoInIsolation
        End Function

        Public Sub DoTransaction(transaction As UndoTransactionImpl)
            Dim _value As LinkedUndoTransactionImpl = Nothing
            If transactionsLinkedTo.TryGetValue(transaction, _value) Then
                _value.Do()
            Else
                transaction.DoInIsolation()
            End If
        End Sub

        Public Sub UndoTransaction(transaction As UndoTransactionImpl)
            Dim _value As LinkedUndoTransactionImpl = Nothing
            If transactionsLinkedTo.TryGetValue(transaction, _value) Then
                _value.Undo()
            Else
                transaction.UndoInIsolation()
            End If
        End Sub

        Public Function AreTransactionsSafeToMerge(transaction1 As UndoTransactionImpl, transaction2 As UndoTransactionImpl) As Boolean
            If currentLinkedUndoTransaction Is Nothing Then
                Return True
            End If

            Dim flag As Boolean = False
            Dim flag2 As Boolean = False

            For Each transaction3 As UndoTransactionImpl In currentLinkedUndoTransaction.Transactions
                If Object.ReferenceEquals(transaction3, transaction1) Then
                    flag = True
                End If

                If Object.ReferenceEquals(transaction3, transaction2) Then
                    flag2 = True
                End If
            Next

            If Not flag OrElse Not flag2 Then
                If Not flag Then
                    Return Not flag2
                End If
                Return False
            End If

            Return True
        End Function

        Public Function GetLinkedUndoTransaction(transaction As UndoTransactionImpl) As LinkedUndoTransactionImpl
            Dim _value As LinkedUndoTransactionImpl = Nothing
            transactionsLinkedTo.TryGetValue(transaction, _value)
            Return _value
        End Function

        Public Sub BreakLinkToUndoTransaction(transaction As UndoTransactionImpl)
            Dim __ As LinkedUndoTransactionImpl = Nothing
            If transactionsLinkedTo.TryGetValue(transaction, __) Then
                transactionsLinkedTo.Remove(transaction)
            End If
        End Sub
    End Class
End Namespace
