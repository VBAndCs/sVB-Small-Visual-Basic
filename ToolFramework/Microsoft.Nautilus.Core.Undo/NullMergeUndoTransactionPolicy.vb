
Namespace Microsoft.Nautilus.Core.Undo
    Public NotInheritable Class NullMergeUndoTransactionPolicy
        Implements IMergeUndoTransactionPolicy

        Private Shared _instance As NullMergeUndoTransactionPolicy

        Public Shared ReadOnly Property Instance As IMergeUndoTransactionPolicy
            Get
                If _instance Is Nothing Then
                    _instance = New NullMergeUndoTransactionPolicy()
                End If

                Return _instance
            End Get
        End Property

        Private Sub New()
        End Sub

        Public Function TestCompatiblePolicy(other As IMergeUndoTransactionPolicy) As Boolean Implements IMergeUndoTransactionPolicy.TestCompatiblePolicy
            Return False
        End Function

        Public Function CanMerge(transaction1 As UndoTransaction, transaction2 As UndoTransaction) As Boolean Implements IMergeUndoTransactionPolicy.CanMerge
            Return False
        End Function

        Public Sub CompleteTransactionMerge(newerTransaction As UndoTransaction, olderTransaction As UndoTransaction, mergedTransaction As UndoTransaction) Implements IMergeUndoTransactionPolicy.CompleteTransactionMerge
            Throw New InvalidOperationException("The NullMergePolicy disallows merging of UndoPrimitives.")
        End Sub

    End Class
End Namespace
