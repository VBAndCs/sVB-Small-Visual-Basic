Imports Microsoft.Nautilus.Text
Imports Microsoft.Nautilus.Text.Projection
Imports Microsoft.Nautilus.Text.Projection.Implementation

Namespace Microsoft.Nautilus.Text.Editor

    Public Class CodeCollapser
        Private _projectionBuffer As IProjectionBuffer
        Private _sourceBuffer As ITextBuffer

        Public Sub New(sourceBuffer As ITextBuffer, contentType As String)
            _sourceBuffer = sourceBuffer
            Dim snapshot = sourceBuffer.CurrentSnapshot
            Dim textSpans As New List(Of ITextSpan) From {
                New TextSpan(snapshot, 0, snapshot.Length, SpanTrackingMode.EdgeExclusive)
            }
            _projectionBuffer = New BufferFactory().CreateProjectionBuffer(Nothing, contentType, textSpans)
        End Sub

        Public ReadOnly Property ProjectionBuffer As IProjectionBuffer
            Get
                Return _projectionBuffer
            End Get
        End Property

        Public Sub CollapseCodeBlock(start As Integer, [end] As Integer)
            Dim snapshot = _sourceBuffer.CurrentSnapshot
            Dim spanToHide As New TextSpan(snapshot, start, [end] - start, SpanTrackingMode.EdgeExclusive)
            Dim spans As New List(Of ITextSpan)()

            ' Add all spans before the code block
            If start > 0 Then
                spans.Add(New TextSpan(snapshot, 0, start, SpanTrackingMode.EdgeExclusive))
            End If

            ' Add all spans after the code block
            If [end] < _sourceBuffer.CurrentSnapshot.Length Then
                spans.Add(New TextSpan(snapshot, [end], snapshot.Length - [end], SpanTrackingMode.EdgeExclusive))
            End If

            _projectionBuffer.ReplaceSpans(0, _projectionBuffer.SourceBuffers.Count, spans)
        End Sub

        Public Sub ExpandCodeBlock(start As Integer, [end] As Integer)
            Dim spans As New List(Of ITextSpan)()
            Dim snapshot = _sourceBuffer.CurrentSnapshot

            ' Add all spans before the code block
            If start > 0 Then
                spans.Add(New TextSpan(snapshot, 0, start, SpanTrackingMode.EdgeExclusive))
            End If

            ' Add the code block itself
            spans.Add(New TextSpan(snapshot, start, [end] - start, SpanTrackingMode.EdgeExclusive))

            ' Add all spans after the code block
            If [end] < _sourceBuffer.CurrentSnapshot.Length Then
                spans.Add(New TextSpan(snapshot, [end], _sourceBuffer.CurrentSnapshot.Length - [end], SpanTrackingMode.EdgeExclusive))
            End If

            _projectionBuffer.ReplaceSpans(0, _projectionBuffer.SourceBuffers.Count, spans)
        End Sub
    End Class

End Namespace