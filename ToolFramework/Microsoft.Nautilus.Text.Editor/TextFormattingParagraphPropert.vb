Imports System.Windows
Imports System.Windows.Media
Imports System.Windows.Media.TextFormatting

Namespace Microsoft.Nautilus.Text.Editor
    Public Class TextFormattingParagraphProperties
        Inherits TextParagraphProperties

        Private _defaultTextRunProperties As TextFormattingRunProperties = TextFormattingRunProperties.CreateTextFormattingRunProperties(New Typeface("Consolas"), 14.0, Colors.Black)

        Public Overrides ReadOnly Property DefaultTextRunProperties As TextRunProperties
            Get
                Return _defaultTextRunProperties
            End Get
        End Property

        Public Overrides ReadOnly Property FirstLineInParagraph As Boolean = False

        Public Overrides ReadOnly Property FlowDirection As FlowDirection = FlowDirection.LeftToRight

        Public Overrides ReadOnly Property TextAlignment As TextAlignment = TextAlignment.Left

        Public Overrides ReadOnly Property Indent As Double = 0.0

        Public Overrides ReadOnly Property LineHeight As Double = 0.0

        Public Overrides ReadOnly Property TextMarkerProperties As TextMarkerProperties = Nothing

        Public Overrides ReadOnly Property TextWrapping As TextWrapping = TextWrapping.Wrap

    End Class
End Namespace
