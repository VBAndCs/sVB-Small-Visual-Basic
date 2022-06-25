Imports Microsoft.Nautilus.Text.Editor

Namespace Microsoft.Nautilus.Text.Adornments
    Public Interface ISimpleProviderViewFactory
        Function GetSimpleProvider(textView As ITextView) As ISimpleProvider
    End Interface
End Namespace
