Namespace Microsoft.Nautilus.Text.Classification
    Public Interface IClassifierAggregatorProvider
        Function GetClassifier(textBuffer As ITextBuffer) As IClassifier
    End Interface
End Namespace
