
Namespace Microsoft.Nautilus.Text
    Public Class TextSpan
        Implements ITextSpan

        Private Class VersionSpan

            Public ReadOnly Property Version As TextVersion

            Public ReadOnly Property Span As Span

            Public Sub New(version As TextVersion, span As Span)
                _Version = version
                _Span = span
            End Sub
        End Class

        Private _version As TextVersion
        Private _span As Span
        Private cachedSpan As VersionSpan

        Public ReadOnly Property TrackingMode As SpanTrackingMode Implements ITextSpan.TrackingMode

        Public ReadOnly Property TextBuffer As ITextBuffer Implements ITextSpan.TextBuffer
            Get
                Return _version.TextBuffer
            End Get
        End Property

        Public Sub New(snapshot As ITextSnapshot, span As Span, trackingMode As SpanTrackingMode)
            If snapshot Is Nothing Then
                Throw New ArgumentNullException("snapshot")
            End If

            If span.End > snapshot.Length Then
                Throw New ArgumentOutOfRangeException("span")
            End If

            _version = snapshot.Version
            _TrackingMode = trackingMode
            _span = span
            cachedSpan = New VersionSpan(_version, _span)
        End Sub

        Public Sub New(snapshot As ITextSnapshot, start As Integer, length As Integer, trackingMode As SpanTrackingMode)
            Me.New(snapshot, New Span(start, length), trackingMode)
        End Sub

        Public Function GetSpan(snapshot As ITextSnapshot) As SnapshotSpan Implements ITextSpan.GetSpan
            If snapshot Is Nothing Then
                Throw New ArgumentNullException("snapshot")
            End If

            If snapshot.TextBuffer IsNot TextBuffer Then
                Throw New ArgumentException("The specified ITextSnapshot doesn't belong to the correct text buffer.")
            End If

            Dim versionSpan1 As VersionSpan = cachedSpan
            Dim span As Span
            If snapshot.Version > versionSpan1.Version Then
                span = MapSpanForwardInTime(versionSpan1.Span, _TrackingMode, snapshot.Version, versionSpan1.Version)
                cachedSpan = New VersionSpan(snapshot.Version, span)

            ElseIf snapshot.Version Is versionSpan1.Version Then
                span = versionSpan1.Span

            ElseIf Not (snapshot.Version > _version) Then
                span = (If((snapshot.Version IsNot _version), MapSpanBackwardsInTime(_span, _TrackingMode, snapshot.Version, _version), _span))

            Else
                span = MapSpanForwardInTime(_span, _TrackingMode, snapshot.Version, _version)
                cachedSpan = New VersionSpan(snapshot.Version, span)
            End If

            Return New SnapshotSpan(snapshot, span)
        End Function

        Public Function GetStartPoint(snapshot As ITextSnapshot) As SnapshotPoint Implements ITextSpan.GetStartPoint
            If snapshot Is Nothing Then
                Throw New ArgumentNullException("snapshot")
            End If

            If snapshot.TextBuffer IsNot TextBuffer Then
                Throw New ArgumentException("The specified ITextSnapshot doesn't belong to the correct text buffer.")
            End If

            Return New SnapshotPoint(snapshot, GetSpan(snapshot).Start)
        End Function

        Public Function GetEndPoint(snapshot As ITextSnapshot) As SnapshotPoint Implements ITextSpan.GetEndPoint
            If snapshot Is Nothing Then
                Throw New ArgumentNullException("snapshot")
            End If

            If snapshot.TextBuffer IsNot TextBuffer Then
                Throw New ArgumentException("The specified ITextSnapshot doesn't belong to the correct text buffer.")
            End If

            Return New SnapshotPoint(snapshot, GetSpan(snapshot).End)
        End Function

        Public Function GetText(snapshot As ITextSnapshot) As String Implements ITextSpan.GetText
            Return GetSpan(snapshot).GetText()
        End Function

        Private Shared Function MapSpanForwardInTime(_span As Span, _trackingMode As SpanTrackingMode, targetVersion As TextVersion, currentVersion As TextVersion) As Span
            Dim num As Integer = TextPoint.MapPositionForwardInTime(_span.Start, If((_trackingMode <> 0), Microsoft.Nautilus.Text.TrackingMode.Negative, Microsoft.Nautilus.Text.TrackingMode.Positive), targetVersion, currentVersion)
            Dim val As Integer = TextPoint.MapPositionForwardInTime(_span.End, If((_trackingMode = SpanTrackingMode.EdgeExclusive), Microsoft.Nautilus.Text.TrackingMode.Negative, Microsoft.Nautilus.Text.TrackingMode.Positive), targetVersion, currentVersion)
            Return Span.FromBounds(num, Math.Max(num, val))
        End Function

        Private Shared Function MapSpanBackwardsInTime(_span As Span, _trackingMode As SpanTrackingMode, targetVersion As TextVersion, currentVersion As TextVersion) As Span
            Dim num As Integer = TextPoint.MapPositionBackwardsInTime(_span.Start, If((_trackingMode <> 0), Microsoft.Nautilus.Text.TrackingMode.Negative, Microsoft.Nautilus.Text.TrackingMode.Positive), targetVersion, currentVersion)
            Dim val As Integer = TextPoint.MapPositionBackwardsInTime(_span.End, If((_trackingMode = SpanTrackingMode.EdgeExclusive), Microsoft.Nautilus.Text.TrackingMode.Negative, Microsoft.Nautilus.Text.TrackingMode.Positive), targetVersion, currentVersion)
            Return Span.FromBounds(num, Math.Max(num, val))
        End Function
    End Class
End Namespace
