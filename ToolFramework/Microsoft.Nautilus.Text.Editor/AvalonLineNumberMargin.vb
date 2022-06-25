Imports System.Collections.Generic
Imports System.Globalization
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input
Imports System.Windows.Media
Imports System.Windows.Media.TextFormatting
Imports Microsoft.Nautilus.Text.Operations

Namespace Microsoft.Nautilus.Text.Editor
    Friend NotInheritable Class AvalonLineNumberMargin
        Inherits Canvas
        Implements IAvalonTextViewMargin, ITextViewMargin

        Private _textView As IAvalonTextView
        Private _editorOperations As IEditorOperations
        Private _textFormatter As TextFormatter
        Private _formatting As TextFormattingRunProperties
        Friend _layoutNeeded As Boolean = True

        Public ReadOnly Property VisualElement As FrameworkElement Implements IAvalonTextViewMargin.VisualElement
            Get
                Return Me
            End Get
        End Property

        Public ReadOnly Property MarginSize As Double Implements ITextViewMargin.MarginSize
            Get
                Return MyBase.ActualWidth
            End Get
        End Property

        Public Property MarginVisible As Boolean Implements ITextViewMargin.MarginVisible
            Get
                Return MyBase.Visibility = Visibility.Visible
            End Get

            Set(value As Boolean)
                If value <> MarginVisible Then
                    If Not value Then
                        UnRegisterEvents()
                        MyBase.Visibility = Visibility.Collapsed
                    Else
                        RegisterEvents()
                        MyBase.Visibility = Visibility.Visible
                    End If
                End If
            End Set
        End Property

        Public Sub New(textView As IAvalonTextView, editorOperations As IEditorOperations)
            If textView Is Nothing Then
                Throw New ArgumentNullException("textView")
            End If

            If editorOperations Is Nothing Then
                Throw New ArgumentNullException("editorOperations")
            End If

            If editorOperations.TextView IsNot textView Then
                Throw New ArgumentException("TextView does not match TextView of the EditorOperations.")
            End If

            _textView = textView
            _editorOperations = editorOperations
            _formatting = TextFormattingRunProperties.CreateTextFormattingRunProperties(New Typeface("Consolas"), 12.0, Color.FromArgb(Byte.MaxValue, 100, 100, 100))
            MyBase.ClipToBounds = True

            MyBase.Background = New LinearGradientBrush(
                  Color.FromArgb(0, 100, 157, 180),
                  Color.FromArgb(Byte.MaxValue, 100, 157, 180), 0.0
                ) With {
                     .GradientStops = New GradientStopCollection From {
                          New GradientStop(Color.FromArgb(48, 100, 157, 180), 0.9999),
                          New GradientStop(Color.FromArgb(Byte.MaxValue, 100, 157, 180), 0.9999)
                     }
                  }

            MyBase.Cursor = Cursors.Arrow
            _textFormatter = TextFormatter.Create()
            Dim textLine1 As TextLine = _textFormatter.FormatLine(New AvalonLineNumberMarginTextSource("88888", _formatting), 0, 50.0, New TextFormattingParagraphProperties, Nothing)
            MyBase.MinWidth = textLine1.Width
            RegisterEvents()
        End Sub

        Private Function SelectLine(mouseLocation As Point, extendSelection As Boolean) As Boolean
            Dim textLineContainingYCoordinate As ITextLine = _textView.FormattedTextLines.GetTextLineContainingYCoordinate(mouseLocation.Y)
            If textLineContainingYCoordinate IsNot Nothing Then
                Dim lineNumberFromPosition As Integer = _textView.TextSnapshot.GetLineNumberFromPosition(textLineContainingYCoordinate.LineSpan.Start)
                _editorOperations.SelectLine(lineNumberFromPosition, extendSelection)
                Return True
            End If
            Return False
        End Function

        Private Sub RegisterEvents()
            AddHandler _textView.LayoutChanged, AddressOf OnEditorLayoutChanged
            AddHandler MyBase.MouseLeftButtonDown, AddressOf OnMouseLeftButtonDown
            AddHandler MyBase.MouseLeftButtonUp, AddressOf OnMouseLeftButtonUp
            AddHandler MyBase.MouseMove, AddressOf OnMouseMove
        End Sub

        Private Sub UnRegisterEvents()
            RemoveHandler _textView.LayoutChanged, AddressOf OnEditorLayoutChanged
            RemoveHandler MyBase.MouseLeftButtonDown, AddressOf OnMouseLeftButtonDown
            RemoveHandler MyBase.MouseLeftButtonUp, AddressOf OnMouseLeftButtonUp
            RemoveHandler MyBase.MouseMove, AddressOf OnMouseMove
        End Sub

        Friend Sub UpdateLineNumbers()
            _layoutNeeded = False
            Dim formattedTextLines1 As IFormattedTextLineCollection = _textView.FormattedTextLines
            If formattedTextLines1.Count <= 0 Then Return

            Dim num As Integer = Integer.MaxValue
            Dim num2 As Integer = Integer.MinValue
            If MyBase.Children.Count > 0 Then
                num = CType(MyBase.Children(0), AvalonLineNumberMarginDrawingVisual).LineNumber
                num2 = CType(MyBase.Children(MyBase.Children.Count - 1), AvalonLineNumberMarginDrawingVisual).LineNumber
            End If

            Dim list1 As New List(Of AvalonLineNumberMarginDrawingVisual)
            Dim num3 As Integer = -1
            For Each item As ITextLine In formattedTextLines1
                Dim num4 As Integer = _textView.TextSnapshot.GetLineNumberFromPosition(item.LineSpan.Start) + 1
                If num4 = num3 Then Continue For

                num3 = num4
                Dim avalonLineNumberMarginDrawingVisual1 As AvalonLineNumberMarginDrawingVisual
                If num <= num4 AndAlso num4 <= num2 Then
                    avalonLineNumberMarginDrawingVisual1 = CType(MyBase.Children(num4 - num), AvalonLineNumberMarginDrawingVisual)
                    Dim point1 As Point = _textView.VisualElement.TranslatePoint(New Point(0.0, item.Top + item.Height), Me)
                    Dim point2 As Point = _textView.VisualElement.TranslatePoint(New Point(0.0, item.Top), Me)
                    avalonLineNumberMarginDrawingVisual1.Update(point1.Y - point2.Y, point2.Y)
                Else
                    Dim text As String = num4.ToString(CultureInfo.InvariantCulture.NumberFormat)
                    If text.Length > "88888".Length Then
                        text = text.Substring(text.Length - "88888".Length)
                    End If
                    Dim textLine1 As TextLine = _textFormatter.FormatLine(New AvalonLineNumberMarginTextSource(text, _formatting), 0, MyBase.MinWidth, New TextFormattingParagraphProperties, Nothing)
                    Dim point3 As Point = _textView.VisualElement.TranslatePoint(New Point(0.0, item.Top + item.Height), Me)
                    Dim point4 As Point = _textView.VisualElement.TranslatePoint(New Point(0.0, item.Top), Me)
                    avalonLineNumberMarginDrawingVisual1 = New AvalonLineNumberMarginDrawingVisual(num4, textLine1, MyBase.MinWidth - textLine1.Width, point3.Y - point4.Y, point4.Y)
                End If
                list1.Add(avalonLineNumberMarginDrawingVisual1)
            Next

            MyBase.Children.Clear()
            For Each item2 As AvalonLineNumberMarginDrawingVisual In list1
                MyBase.Children.Add(item2)
            Next
        End Sub

        Friend Sub OnEditorLayoutChanged(sender As Object, e As EventArgs)
            _layoutNeeded = True
            InvalidateVisual()
        End Sub

        Private Overloads Sub OnMouseMove(sender As Object, e As MouseEventArgs)
            If e.LeftButton <> MouseButtonState.Pressed Then Return

            Dim position As Point = e.GetPosition(Me)
            If Not SelectLine(position, extendSelection:=True) Then
                If position.Y <= 0.0 Then
                    _editorOperations.MoveLineUp([select]:=True)
                ElseIf position.Y > _textView.ViewportHeight Then
                    _editorOperations.MoveLineDown([select]:=True)
                End If
                _textView.Caret.EnsureVisible()
            End If
        End Sub

        Private Overloads Sub OnMouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs)
            CaptureMouse()
            Dim extendSelection As Boolean = (Keyboard.Modifiers And ModifierKeys.Shift) = ModifierKeys.Shift
            SelectLine(e.GetPosition(Me), extendSelection)
        End Sub

        Private Overloads Sub OnMouseLeftButtonUp(sender As Object, e As MouseButtonEventArgs)
            ReleaseMouseCapture()
        End Sub

        Protected Overrides Sub OnRender(dc As DrawingContext)
            If _layoutNeeded Then
                UpdateLineNumbers()
            End If
            MyBase.OnRender(dc)
        End Sub
    End Class
End Namespace
