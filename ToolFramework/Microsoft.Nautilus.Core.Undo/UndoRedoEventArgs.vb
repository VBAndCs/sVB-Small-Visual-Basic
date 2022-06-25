
Namespace Microsoft.Nautilus.Core.Undo
    Public Class UndoRedoEventArgs
        Inherits EventArgs

        Public ReadOnly Property Transaction As UndoTransaction

        Public ReadOnly Property State As UndoHistoryState

        Public Sub New(state As UndoHistoryState, transaction As UndoTransaction)
            _State = state
            _Transaction = transaction
        End Sub
    End Class
End Namespace
