Imports Microsoft.Nautilus.Core.Undo

Namespace Microsoft.Nautilus.Text.Operations
    Friend Class TextTransactionMergePolicy
        Implements IMergeUndoTransactionPolicy

        Public Function CanMerge(transaction1 As UndoTransaction, transaction2 As UndoTransaction) As Boolean Implements IMergeUndoTransactionPolicy.CanMerge
            If transaction1 Is Nothing Then
                Throw New ArgumentNullException("transaction1")
            End If

            If transaction2 Is Nothing Then
                Throw New ArgumentNullException("transaction2")
            End If

            Return transaction1.Description = transaction2.Description
        End Function

        Public Sub CompleteTransactionMerge(newerTransaction As UndoTransaction, olderTransaction As UndoTransaction, mergedTransaction As UndoTransaction) Implements IMergeUndoTransactionPolicy.CompleteTransactionMerge
            If newerTransaction Is Nothing Then
                Throw New ArgumentNullException("newerTransaction")
            End If

            If olderTransaction Is Nothing Then
                Throw New ArgumentNullException("olderTransaction")
            End If

            If mergedTransaction Is Nothing Then
                Throw New ArgumentNullException("mergedTransaction")
            End If
        End Sub

        Public Function TestCompatiblePolicy(other As IMergeUndoTransactionPolicy) As Boolean Implements IMergeUndoTransactionPolicy.TestCompatiblePolicy
            If other Is Nothing Then
                Throw New ArgumentNullException("other")
            End If

            Return [GetType]() = other.GetType()
        End Function
    End Class
End Namespace
