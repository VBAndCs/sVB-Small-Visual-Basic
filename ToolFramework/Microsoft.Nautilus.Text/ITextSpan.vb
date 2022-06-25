Namespace Microsoft.Nautilus.Text
    Public Interface ITextSpan
        ReadOnly Property TextBuffer As ITextBuffer

        ReadOnly Property TrackingMode As SpanTrackingMode

        Function GetSpan(snapshot As ITextSnapshot) As SnapshotSpan

        Function GetText(snapshot As ITextSnapshot) As String

        Function GetStartPoint(snapshot As ITextSnapshot) As SnapshotPoint

        Function GetEndPoint(snapshot As ITextSnapshot) As SnapshotPoint
    End Interface
End Namespace
