Imports Microsoft.Nautilus.Text.Editor

Namespace Microsoft.Nautilus.Text.Adornments
    Public Interface IToolTipProviderFactory
        Function GetToolTipProvider(textView As ITextView) As IToolTipProvider
    End Interface
End Namespace
