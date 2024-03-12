Imports System.Windows.Media
Imports Microsoft.Nautilus.Text
Imports Microsoft.Nautilus.Text.AdornmentSystem

Namespace Microsoft.Windows.Controls
    Public Class StatementMarker
        Implements IAdornment

        Public LineNumber As Integer

        Dim _markerColor As Color
        Friend RefreshMarker As Action(Of StatementMarker)

        Public Property MarkerColor As Color
            Get
                Return _markerColor
            End Get

            Set(value As Color)
                _markerColor = value
                RefreshMarker(Me)
            End Set
        End Property

        Public Property Span As ITextSpan Implements IAdornment.Span

        Public Sub New(span As ITextSpan, lineNumber As Integer, markerColor As Color)
            If span Is Nothing Then
                Throw New ArgumentNullException("span")
            End If

            _Span = span
            _MarkerColor = markerColor
            Me.LineNumber = lineNumber
        End Sub

        Public Sub New(line As ITextSnapshotLine, markerColor As Color)
            Me.New(New TextSpan(
                                line.TextSnapshot,
                                line.Start,
                                line.Length,
                                SpanTrackingMode.EdgeExclusive),
                            line.LineNumber,
                            markerColor)
        End Sub
    End Class
End Namespace
