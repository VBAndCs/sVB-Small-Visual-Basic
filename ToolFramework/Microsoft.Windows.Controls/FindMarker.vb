Imports System.Windows.Media
Imports Microsoft.Nautilus.Text
Imports Microsoft.Nautilus.Text.AdornmentSystem

Namespace Microsoft.Windows.Controls
    Public Class FindMarker
        Implements IAdornment

        Public Property Brush As Brush

        Public Property MarkerColor As Color

        Public Property Pen As Pen

        Public Property Span As ITextSpan Implements IAdornment.Span

        Public Sub New(span As ITextSpan, markerColor As Color)
            If span Is Nothing Then
                Throw New ArgumentNullException("span")
            End If

            _Span = span
            _MarkerColor = markerColor
            _Pen = New Pen(New SolidColorBrush(markerColor), 1.0)
            _Brush = New SolidColorBrush(Color.FromArgb(64, markerColor.R, markerColor.G, markerColor.B))
        End Sub
    End Class
End Namespace
