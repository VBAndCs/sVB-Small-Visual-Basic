Imports System.Globalization

Namespace Microsoft.Nautilus.Text
    Public Class ReadOnlyRegionHandle
        Implements IReadOnlyRegionHandle

        Private _readOnlyRegionManager As ReadOnlyRegionManager

        Public ReadOnly Property Span As ITextSpan Implements IReadOnlyRegionHandle.Span

        Public ReadOnly Property EdgeInsertionMode As EdgeInsertionMode Implements IReadOnlyRegionHandle.EdgeInsertionMode

        Friend Sub New(textSnapshot As ITextSnapshot, textSpan As Span, trackingMode As SpanTrackingMode, edgeInsertionMode1 As EdgeInsertionMode, readOnlySpanManager As ReadOnlyRegionManager)
            _EdgeInsertionMode = edgeInsertionMode1
            _readOnlyRegionManager = readOnlySpanManager
            _Span = New TextSpan(textSnapshot, textSpan, trackingMode)
        End Sub

        Public Sub Remove() Implements IReadOnlyRegionHandle.Remove
            _readOnlyRegionManager.RemoveReadOnlyRegionHandle(Me)
        End Sub

        Public Overrides Function ToString() As String
            Dim span1 = _Span.GetSpan(_Span.TextBuffer.CurrentSnapshot)
            Return String.Format(CultureInfo.InvariantCulture, "RO: {2}{0}..{1}{3}", span1.Start, span1.End, If((_EdgeInsertionMode = EdgeInsertionMode.Deny), "[", "("), If((_EdgeInsertionMode = EdgeInsertionMode.Deny), "]", ")"))
        End Function
    End Class
End Namespace
