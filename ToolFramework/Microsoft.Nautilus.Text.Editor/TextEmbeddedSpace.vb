Imports System.Windows
Imports System.Windows.Media
Imports System.Windows.Media.TextFormatting

Namespace Microsoft.Nautilus.Text.Editor
    Friend Class TextEmbeddedSpace
        Inherits TextEmbeddedObject

        Private _size As Size
        Private ReadOnly _defaultTextRunProperties As TextRunProperties = TextFormattingRunProperties.CreateTextFormattingRunProperties(New Typeface(SystemFonts.CaptionFontFamily, SystemFonts.CaptionFontStyle, SystemFonts.CaptionFontWeight, FontStretches.Normal), SystemFonts.CaptionFontSize, SystemColors.WindowColor)

        Public Overrides ReadOnly Property BreakAfter As LineBreakCondition
            Get
                Return LineBreakCondition.BreakPossible
            End Get
        End Property

        Public Overrides ReadOnly Property BreakBefore As LineBreakCondition
            Get
                Return LineBreakCondition.BreakPossible
            End Get
        End Property

        Public Overrides ReadOnly Property HasFixedSize As Boolean = True

        Public Overrides ReadOnly Property CharacterBufferReference As CharacterBufferReference
            Get
                Return New CharacterBufferReference(" ", 1)
            End Get
        End Property

        Public Overrides ReadOnly Property Length As Integer = 1

        Public Overrides ReadOnly Property Properties As TextRunProperties
            Get
                Return _defaultTextRunProperties
            End Get
        End Property

        Public Sub New(size As Size)
            _size = size
        End Sub

        Public Overrides Function ComputeBoundingBox(rightToLeft As Boolean, sideways As Boolean) As Rect
            Return New Rect(0.0, 0.0, _size.Width, _size.Height)
        End Function

        Public Overrides Sub Draw(drawingContext1 As DrawingContext, origin As Point, rightToLeft As Boolean, sideways As Boolean)
        End Sub

        Public Overrides Function Format(remainingParagraphWidth As Double) As TextEmbeddedObjectMetrics
            Return New TextEmbeddedObjectMetrics(_size.Width, _size.Height, _size.Height)
        End Function

    End Class
End Namespace
