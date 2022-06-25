Imports System.Collections.Generic

Namespace Microsoft.Nautilus.Text.Projection
    Public Interface IProjectionSnapshot
        Inherits ITextSnapshot
        ReadOnly Property SpanCount As Integer

        ReadOnly Property SourceSnapshots As IList(Of ITextSnapshot)

        Function GetMatchingSnapshot(textBuffer As ITextBuffer) As ITextSnapshot

        Function GetSourceSpans(startSpanIndex As Integer, count As Integer) As IList(Of SnapshotSpan)

        Function GetSourceSpans() As IList(Of SnapshotSpan)

        Function MapToSourceSnapshot(position As Integer) As SnapshotPoint

        Function MapFromSourceSnapshot(point As SnapshotPoint) As SnapshotPoint?

        Function MapToSourceSnapshots(span1 As Span) As IList(Of SnapshotSpan)

        Function MapFromSourceSnapshot(span1 As SnapshotSpan) As IList(Of Span)
    End Interface
End Namespace
