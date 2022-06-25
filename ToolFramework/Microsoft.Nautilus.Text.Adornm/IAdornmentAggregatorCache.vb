Imports Microsoft.Nautilus.Text.Editor

Namespace Microsoft.Nautilus.Text.AdornmentSystem
    Public Interface IAdornmentAggregatorCache
        Function GetAdornmentProvider(textView As ITextView) As IAdornmentProvider
    End Interface
End Namespace
