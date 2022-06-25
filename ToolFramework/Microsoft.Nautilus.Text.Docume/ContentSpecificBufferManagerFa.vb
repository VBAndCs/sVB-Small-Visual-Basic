Namespace Microsoft.Nautilus.Text.Document
    Public MustInherit Class ContentSpecificBufferManagerFactory
        Public MustOverride Function CreateBufferManager(textBuffer As ITextBuffer) As IBufferManager
    End Class
End Namespace
