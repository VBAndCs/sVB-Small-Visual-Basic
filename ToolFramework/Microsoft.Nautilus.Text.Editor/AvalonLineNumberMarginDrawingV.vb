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
        Private _ShowBreakpoint As Boolean
        Public ReadOnly Property VerticalOffset As Double

        Public ReadOnly Property LineNumber As Integer

        Public Property ShowBreakpoint As Boolean
            Get
                Return _ShowBreakpoint
            End Get

            Set
                If _ShowBreakpoint <> Value Then
                    _ShowBreakpoint = Value
                    InvalidateVisual()
                End If
            End Set
        End Property

        Private Sub CalculateVerticalOffset()
            If _viewLineHeight > _textLine.Height Then
                _VerticalOffset = _viewLineTop + (_viewLineHeight - _textLine.Height) * 0.5 + 1.0
            Else
                _VerticalOffset = _viewLineTop + 1.0
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
            Dim verticalOffset = _VerticalOffset
            _viewLineHeight = viewLineHeight
            _viewLineTop = viewLineTop
            CalculateVerticalOffset()

            If verticalOffset <> _VerticalOffset Then
                InvalidateVisual()
            End If
        End Sub

        Protected Overrides Sub OnRender(drawingContext As DrawingContext)
            If _ShowBreakpoint Then
                Dim w = (CType(Me.Parent, Controls.Canvas).ActualWidth) / 2
                Dim h = _textLine.Height / 2
                Dim center = New Point(w, _VerticalOffset + h)
                drawingContext.DrawEllipse(Brushes.DarkRed, Nothing, center, h, h)
            End If

            Dim origin = New Point(_horizontalOffset, _VerticalOffset)
            _textLine.Draw(drawingContext, origin, InvertAxes.None)

        End Sub
    End Class
End Namespace
