Imports System.Collections.Generic

Namespace Microsoft.Nautilus.Text.Classification
    Public Interface IClassifier
        Event ClassificationChanged As EventHandler(Of ClassificationChangedEventArgs)

        Function GetClassificationSpans(textSpan As SnapshotSpan) As IList(Of ClassificationSpan)
    End Interface
End Namespace
