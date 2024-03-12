Imports System.Collections.Generic
Imports System.Globalization
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input
Imports System.Windows.Media
Imports System.Windows.Media.TextFormatting
Imports Microsoft.Nautilus.Text.Operations

Namespace Microsoft.Nautilus.Text.Editor
    Public NotInheritable Class AvalonLineNumberMargin
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
            Dim textLine = _textFormatter.FormatLine(New AvalonLineNumberMarginTextSource("8888", _formatting), 0, 50.0, New TextFormattingParagraphProperties, Nothing)
            MyBase.MinWidth = textLine.Width
            RegisterEvents()
        End Sub

        Private Function SelectLine(mouseLocation As Point, extendSelection As Boolean) As Boolean
            Dim textLine = _textView.FormattedTextLines.GetTextLineContainingYCoordinate(mouseLocation.Y)
            If textLine Is Nothing Then Return False

            Dim lineNumber = _textView.TextSnapshot.GetLineNumberFromPosition(textLine.LineSpan.Start)
            _editorOperations.SelectLine(lineNumber, extendSelection)
            Return True
        End Function

        Public Event LineBreakpointChanged(ByRef lineNumber As Integer, ByRef showBreeakPoint As Boolean)

        Private Sub ToogleBreakpoint(mouseLocation As Point)
            Dim textLine = _textView.FormattedTextLines.GetTextLineContainingYCoordinate(mouseLocation.Y)
            If textLine Is Nothing Then Return

            Dim lineNumber = _textView.TextSnapshot.GetLineNumberFromPosition(textLine.LineSpan.Start)
            Dim showBreeakPoint As Boolean
            RaiseEvent LineBreakpointChanged(lineNumber, showBreeakPoint)
            If lineNumber > -1 Then DrawBreakpoint(lineNumber, showBreeakPoint)
        End Sub

        Public Sub ClearBreakpoint()
            For Each visual As AvalonLineNumberMarginDrawingVisual In MyBase.Children
                visual.ShowBreakpoint = False
            Next
        End Sub

        Public Sub DrawBreakpoint(lineNumber As Integer, showBreeakPoint As Boolean)
            lineNumber += 1
            For Each visual As AvalonLineNumberMarginDrawingVisual In MyBase.Children
                If visual.LineNumber = lineNumber Then
                    visual.ShowBreakpoint = showBreeakPoint
                    Return
                End If
            Next
        End Sub


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
            Dim formattedTextLines = _textView.FormattedTextLines
            If formattedTextLines.Count <= 0 Then Return

            Dim start As Integer = Integer.MaxValue
            Dim [end] As Integer = Integer.MinValue
            If MyBase.Children.Count > 0 Then
                start = CType(MyBase.Children(0), AvalonLineNumberMarginDrawingVisual).LineNumber
                [end] = CType(MyBase.Children(MyBase.Children.Count - 1), AvalonLineNumberMarginDrawingVisual).LineNumber
            End If

            Dim visuals As New List(Of AvalonLineNumberMarginDrawingVisual)
            Dim prevLineNum As Integer = -1
            For Each formattedTextLine As ITextLine In formattedTextLines
                Dim lineNum = _textView.TextSnapshot.GetLineNumberFromPosition(formattedTextLine.LineSpan.Start) + 1
                If lineNum = prevLineNum Then Continue For

                prevLineNum = lineNum
                Dim lineMarginVisual As AvalonLineNumberMarginDrawingVisual
                If start <= lineNum AndAlso lineNum <= [end] Then
                    lineMarginVisual = CType(MyBase.Children(lineNum - start), AvalonLineNumberMarginDrawingVisual)
                    Dim bottom = _textView.VisualElement.TranslatePoint(New Point(0.0, formattedTextLine.Top + formattedTextLine.Height), Me)
                    Dim top = _textView.VisualElement.TranslatePoint(New Point(0.0, formattedTextLine.Top), Me)
                    lineMarginVisual.Update(bottom.Y - top.Y, top.Y)
                Else
                    Dim text = lineNum.ToString(CultureInfo.InvariantCulture.NumberFormat)
                    If text.Length > "8888".Length Then
                        text = text.Substring(text.Length - "8888".Length)
                    End If
                    Dim textLine = _textFormatter.FormatLine(New AvalonLineNumberMarginTextSource(text, _formatting), 0, MyBase.MinWidth, New TextFormattingParagraphProperties, Nothing)
                    Dim bottom As Point = _textView.VisualElement.TranslatePoint(New Point(0.0, formattedTextLine.Top + formattedTextLine.Height), Me)
                    Dim top As Point = _textView.VisualElement.TranslatePoint(New Point(0.0, formattedTextLine.Top), Me)
                    lineMarginVisual = New AvalonLineNumberMarginDrawingVisual(lineNum, textLine, MyBase.MinWidth - textLine.Width, bottom.Y - top.Y, top.Y)
                End If
                visuals.Add(lineMarginVisual)
            Next

            MyBase.Children.Clear()
            For Each lineMarginVisual As AvalonLineNumberMarginDrawingVisual In visuals
                MyBase.Children.Add(lineMarginVisual)
            Next
        End Sub

        Friend Sub OnEditorLayoutChanged(sender As Object, e As EventArgs)
            _layoutNeeded = True
            InvalidateVisual()
        End Sub

        Private Overloads Sub OnMouseMove(sender As Object, e As MouseEventArgs)
            If e.LeftButton <> MouseButtonState.Pressed Then Return

            Dim position = e.GetPosition(Me)
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
            If e.ClickCount > 1 Then
                ToogleBreakpoint(e.GetPosition(Me))
                e.Handled = True
            Else
                CaptureMouse()
                Dim extendSelection = (Keyboard.Modifiers And ModifierKeys.Shift) = ModifierKeys.Shift
                SelectLine(e.GetPosition(Me), extendSelection)
            End If
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
