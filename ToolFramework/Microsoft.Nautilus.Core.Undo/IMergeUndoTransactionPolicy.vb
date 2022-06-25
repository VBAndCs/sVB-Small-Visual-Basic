Namespace Microsoft.Nautilus.Core.Undo
    Public Interface IMergeUndoTransactionPolicy
        Function TestCompatiblePolicy(other As IMergeUndoTransactionPolicy) As Boolean

        Function CanMerge(transaction1 As UndoTransaction, transaction2 As UndoTransaction) As Boolean

        Sub CompleteTransactionMerge(newerTransaction As UndoTransaction, olderTransaction As UndoTransaction, mergedTransaction As UndoTransaction)
    End Interface
End Namespace
