Imports Microsoft.Nautilus.Core.Undo

Namespace Microsoft.Nautilus.Text.Operations
    Public Interface ITextBufferUndoManager
        ReadOnly Property TextBuffer As ITextBuffer

        ReadOnly Property TextBufferUndoHistory As UndoHistory

        Sub UnregisterUndoHistory()
    End Interface
End Namespace
