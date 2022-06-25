Namespace Microsoft.Nautilus.Text.Operations
    Public Interface ITextBufferUndoManagerProvider
        Function GetTextBufferUndoManager(textBuffer As ITextBuffer) As ITextBufferUndoManager
    End Interface
End Namespace
