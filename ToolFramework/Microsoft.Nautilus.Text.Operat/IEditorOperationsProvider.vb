Imports Microsoft.Nautilus.Text.Editor

Namespace Microsoft.Nautilus.Text.Operations
    Public Interface IEditorOperationsProvider
        Function GetEditorOperations(textView As ITextView) As IEditorOperations
    End Interface
End Namespace
