Imports System.Windows.Media
Imports Microsoft.Nautilus.Text
Imports Microsoft.Nautilus.Text.AdornmentSystem

Namespace Microsoft.Windows.Controls
    Public Class StatementMarker
        Implements IAdornment

        Public Property MarkerColor As Color

        Public Property Span As ITextSpan Implements IAdornment.Span

        Public Sub New(span As ITextSpan, markerColor As Color)
            If span Is Nothing Then
                Throw New ArgumentNullException("span")
            End If

            _Span = span
            _MarkerColor = markerColor
        End Sub
    End Class
End Namespace
