Namespace Microsoft.Nautilus.Text.Projection
    Public Interface IBufferGraph
        ReadOnly Property TopBuffer As ITextBuffer

        Function MapDownToBuffer(topPosition As Integer, targetBuffer As ITextBuffer) As SnapshotPoint

        Function MapDownToBuffer(topSpan As Span, targetBuffer As ITextBuffer) As NormalizedSpanCollection

        Function MapUpToTop(point As SnapshotPoint) As SnapshotPoint?

        Function MapUpToTop(span1 As SnapshotSpan) As NormalizedSpanCollection
    End Interface
End Namespace
