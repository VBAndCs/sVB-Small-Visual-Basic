Imports System.Windows
Imports System.Windows.Media
Imports System.Windows.Media.TextFormatting

Namespace Microsoft.Nautilus.Text.Editor

    Friend Class AvalonLineNumberMarginDrawingVisual
        Inherits FrameworkElement

        Private _textLine As TextLine
        Private _horizontalOffset As Double
        Private _viewLineHeight As Double
        Private _viewLineTop As Double

        Public ReadOnly Property VerticalOffset As Double

        Public ReadOnly Property LineNumber As Integer

        Private Sub CalculateVerticalOffset()
            If _viewLineHeight > _textLine.Height Then
                _verticalOffset = _viewLineTop + (_viewLineHeight - _textLine.Height) * 0.5 + 1.0
            Else
                _verticalOffset = _viewLineTop + 1.0
            End If
        End Sub

        Public Sub New(lineNumber As Integer, textLine As TextLine, horizontalOffset As Double, viewLineHeight As Double, viewLineTop As Double)
            If lineNumber < 1 Then
                Throw New ArgumentOutOfRangeException("lineNumber")
            End If

            If textLine Is Nothing Then
                Throw New ArgumentNullException("textLine")
            End If

            If horizontalOffset < 0.0 Then
                Throw New ArgumentOutOfRangeException("horizontalOffset")
            End If

            If viewLineHeight < 0.0 Then
                Throw New ArgumentOutOfRangeException("viewLineHeight")
            End If

            _LineNumber = lineNumber
            _textLine = textLine
            _horizontalOffset = horizontalOffset
            _viewLineHeight = viewLineHeight
            _viewLineTop = viewLineTop
            CalculateVerticalOffset()
        End Sub

        Public Sub Update(viewLineHeight As Double, viewLineTop As Double)
            Dim verticalOffset1 As Double = _verticalOffset
            _viewLineHeight = viewLineHeight
            _viewLineTop = viewLineTop
            CalculateVerticalOffset()

            If verticalOffset1 <> _verticalOffset Then
                InvalidateVisual()
            End If
        End Sub

        Protected Overrides Sub OnRender(drawingContext As DrawingContext)
            _textLine.Draw(drawingContext, New Point(_horizontalOffset, _VerticalOffset), InvertAxes.None)
        End Sub
    End Class
End Namespace
