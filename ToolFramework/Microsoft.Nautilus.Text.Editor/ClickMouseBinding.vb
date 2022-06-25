Imports System.Windows
Imports System.Windows.Input
Imports Microsoft.Nautilus.Text.Operations

Namespace Microsoft.Nautilus.Text.Editor
    Friend Class ClickMouseBinding
        Implements IMouseBinding

        Private _textView As IAvalonTextView
        Private _editorOperations As IEditorOperations
        Private _selectedWordSpan As ITextSpan
        Private _selectedWordSpanIsReversed As Boolean
        Private _textStructureNavigator As ITextStructureNavigator
        Private _caretAtEnd As Boolean

        Friend Sub New(textView As IAvalonTextView, editorOperationsProvider As IEditorOperationsProvider, textStructureNavigatorFactory As ITextStructureNavigatorFactory)
            If textView Is Nothing Then
                Throw New ArgumentNullException("textView")
            End If

            If editorOperationsProvider Is Nothing Then
                Throw New ArgumentNullException("editorOperationsProvider")
            End If

            If textStructureNavigatorFactory Is Nothing Then
                Throw New ArgumentNullException("textStructureNavigatorFactory")
            End If

            _textView = textView
            _editorOperations = editorOperationsProvider.GetEditorOperations(textView)
            _textStructureNavigator = textStructureNavigatorFactory.CreateTextStructureNavigator(textView.TextBuffer)
        End Sub

        Public Sub PreprocessMouseLeftButtonDown(e As MouseButtonEventArgs) Implements IMouseBinding.PreprocessMouseLeftButtonDown
            If e Is Nothing Then
                Throw New ArgumentNullException("e")
            End If
            Dim shift1 As Boolean = (Keyboard.Modifiers And ModifierKeys.Shift) = ModifierKeys.Shift
            Dim control1 As Boolean = (Keyboard.Modifiers And ModifierKeys.Control) = ModifierKeys.Control
            If e.ClickCount = 1 Then
                e.Handled = HandleSingleClick(e, shift1, control1)
            ElseIf e.ClickCount = 2 Then
                e.Handled = HandleDoubleClick()
            End If
            If Not e.Handled Then
                _selectedWordSpan = Nothing
            End If
        End Sub

        Public Sub PreprocessMouseLeftButtonUp(e As MouseButtonEventArgs) Implements IMouseBinding.PreprocessMouseLeftButtonUp
            _selectedWordSpan = Nothing
        End Sub

        Public Sub PreprocessMouseMove(e As MouseEventArgs) Implements IMouseBinding.PreprocessMouseMove
            If e Is Nothing Then
                Throw New ArgumentNullException("e")
            End If
            If _selectedWordSpan Is Nothing Then
                Return
            End If
            _textView.VisualElement.CaptureMouse()
            Dim position As Point = e.GetPosition(_textView.VisualElement)
            position.X += _textView.ViewportLeft
            Dim formattedTextLines1 As IFormattedTextLineCollection = _textView.FormattedTextLines
            Dim textLine As ITextLine
            Dim index As Integer

            If formattedTextLines1.FindTextLineIndexContainingYCoordinate(position.Y, index) Then
                textLine = formattedTextLines1(index)
            ElseIf index = -1 Then
                textLine = formattedTextLines1(0)
                _textView.ViewScroller.ScrollViewportVerticallyByLine(ScrollDirection.Up)
                index = formattedTextLines1.GetIndexOfTextLine(textLine)
                If index > 0 Then
                    textLine = formattedTextLines1(index - 1)
                End If
            Else
                textLine = formattedTextLines1(formattedTextLines1.Count - 1)
                _textView.ViewScroller.ScrollViewportVerticallyByLine(ScrollDirection.Down)
                index = formattedTextLines1.GetIndexOfTextLine(textLine)
                If index < formattedTextLines1.Count - 1 Then
                    textLine = formattedTextLines1(index + 1)
                End If
            End If

            Dim caretPosition As ICaretPosition = textLine.MoveCaretToLocation(position.X)
            caretPosition = ExtendDoubleClickSelection(caretPosition)
            UpdateCaret()
            e.Handled = True
        End Sub

        Public Sub PreprocessMouseEnter(e As MouseEventArgs) Implements IMouseBinding.PreprocessMouseEnter
        End Sub

        Public Sub PreprocessMouseLeave(e As MouseEventArgs) Implements IMouseBinding.PreprocessMouseLeave
        End Sub

        Public Sub PreprocessMouseRightButtonDown(e As MouseButtonEventArgs) Implements IMouseBinding.PreprocessMouseRightButtonDown
        End Sub

        Public Sub PreprocessMouseRightButtonUp(e As MouseButtonEventArgs) Implements IMouseBinding.PreprocessMouseRightButtonUp
        End Sub

        Public Sub PreprocessMouseWheel(e As MouseWheelEventArgs) Implements IMouseBinding.PreprocessMouseWheel
        End Sub

        Public Sub PostprocessMouseEnter(e As MouseEventArgs) Implements IMouseBinding.PostprocessMouseEnter
        End Sub

        Public Sub PostprocessMouseLeave(e As MouseEventArgs) Implements IMouseBinding.PostprocessMouseLeave
        End Sub

        Public Sub PostprocessMouseLeftButtonUp(e As MouseButtonEventArgs) Implements IMouseBinding.PostprocessMouseLeftButtonUp
        End Sub

        Public Sub PostprocessMouseMove(e As MouseEventArgs) Implements IMouseBinding.PostprocessMouseMove
        End Sub

        Public Sub PostprocessMouseLeftButtonDown(e As MouseButtonEventArgs) Implements IMouseBinding.PostprocessMouseLeftButtonDown
        End Sub

        Public Sub PostprocessMouseRightButtonDown(e As MouseButtonEventArgs) Implements IMouseBinding.PostprocessMouseRightButtonDown
        End Sub

        Public Sub PostprocessMouseRightButtonUp(e As MouseButtonEventArgs) Implements IMouseBinding.PostprocessMouseRightButtonUp
        End Sub

        Public Sub PostprocessMouseWheel(e As MouseWheelEventArgs) Implements IMouseBinding.PostprocessMouseWheel
        End Sub

        Private Function ExtendDoubleClickSelection(caretPosition As ICaretPosition) As ICaretPosition
            Dim extentOfCurrentWord As TextExtent = GetExtentOfCurrentWord(caretPosition.TextInsertionIndex)
            Dim num As Integer = extentOfCurrentWord.Start
            Dim span1 As Span = _selectedWordSpan.GetSpan(_textView.TextSnapshot)
            Dim num2 As Integer
            If Not _selectedWordSpanIsReversed Then
                num2 = (If((caretPosition.TextInsertionIndex > span1.Start), span1.Start, extentOfCurrentWord.Start))
                num = (If((caretPosition.TextInsertionIndex > span1.End), (extentOfCurrentWord.Start + extentOfCurrentWord.Length), span1.End))
            Else
                num2 = (If((caretPosition.TextInsertionIndex < span1.End), span1.Start, (extentOfCurrentWord.Start + extentOfCurrentWord.Length)))
                If caretPosition.TextInsertionIndex >= span1.Start Then
                    num = span1.End
                End If
            End If
            If CInt(CLng(Fix(_textView.Selection.ActiveSpan.GetStartPoint(_textView.TextSnapshot))) Mod Integer.MaxValue) < span1.Start Then
                _caretAtEnd = False
            ElseIf CInt(CLng(Fix(_textView.Selection.ActiveSpan.GetEndPoint(_textView.TextSnapshot))) Mod Integer.MaxValue) > span1.End Then
                _caretAtEnd = True
            End If
            Dim result As ICaretPosition = (If((Not _caretAtEnd), _textView.Caret.MoveTo(num2, CaretPlacement.LeadingEdgeOfCharacter), _textView.Caret.MoveTo(num, CaretPlacement.LeadingEdgeOfCharacter)))
            If num2 < num Then
                _textView.Selection.ActiveSpan = _textView.TextSnapshot.CreateTextSpan(Span.FromBounds(num2, num), SpanTrackingMode.EdgeExclusive)
                _textView.Selection.IsActiveSpanReversed = False
            Else
                _textView.Selection.ActiveSpan = _textView.TextSnapshot.CreateTextSpan(Span.FromBounds(num, num2), SpanTrackingMode.EdgeExclusive)
                _textView.Selection.IsActiveSpanReversed = True
            End If
            Return result
        End Function
        Friend Function GetExtentOfCurrentWord(currentPosition As Integer) As TextExtent
            Dim extentOfWord As TextExtent = _textStructureNavigator.GetExtentOfWord(New SnapshotPoint(_textView.TextSnapshot, currentPosition))
            If extentOfWord.IsSignificant Then
                Return extentOfWord
            End If
            Dim lineFromPosition As ITextSnapshotLine = _textView.TextSnapshot.GetLineFromPosition(currentPosition)
            Dim start1 As Integer = lineFromPosition.Start
            Dim [end] As Integer = lineFromPosition.End
            Dim num As Integer = Math.Max(start1, extentOfWord.Start)
            Dim num2 As Integer = Math.Min([end], extentOfWord.Start + extentOfWord.Length)
            Dim num3 As Integer = num2 - num
            Select Case num3
                Case 0
                    If num = start1 Then
                        Return New TextExtent(currentPosition, 0, isSignificant:=False)
                    End If
                    Return GetExtentOfCurrentWord(num - 1)

                Case 1
                    If num = start1 Then
                        Return New TextExtent(num, 1, isSignificant:=False)
                    End If
                    Return GetExtentOfCurrentWord(num - 1)

                Case Else
                    Return New TextExtent(num, num3, isSignificant:=False)
            End Select
        End Function

        Private Sub UpdateCaret()
            _textView.Caret.EnsureVisible()
            _textView.Caret.CapturePreferredBounds()
        End Sub

        Private Sub SelectWordUnderCaret()
            _editorOperations.SelectCurrentWord()
            _selectedWordSpan = _textView.Selection.ActiveSpan
            _selectedWordSpanIsReversed = _textView.Selection.IsActiveSpanReversed
            _caretAtEnd = True
        End Sub
        Friend Function HandleDoubleClick() As Boolean
            SelectWordUnderCaret()
            Return True
        End Function
        Friend Function HandleSingleClick(e As MouseButtonEventArgs, shift1 As Boolean, control1 As Boolean) As Boolean
            Dim result As Boolean = False
            _textView.VisualElement.CaptureMouse()
            Dim position As Point = e.GetPosition(_textView.VisualElement)
            position.X += _textView.ViewportLeft
            Dim textLineContainingYCoordinate As ITextLine = _textView.FormattedTextLines.GetTextLineContainingYCoordinate(position.Y)

            If textLineContainingYCoordinate IsNot Nothing AndAlso control1 Then
                Dim caretPosition As ICaretPosition = textLineContainingYCoordinate.MoveCaretToLocation(position.X)
                If shift1 Then
                    Dim extentOfCurrentWord As TextExtent = GetExtentOfCurrentWord(caretPosition.TextInsertionIndex)
                    Dim flag As Boolean = False
                    Dim num As Integer

                    If _textView.Selection.IsEmpty Then
                        flag = _textView.Caret.Position.TextInsertionIndex <= caretPosition.TextInsertionIndex
                        num = _textView.Caret.Position.TextInsertionIndex
                    Else
                        num = (If((Not _textView.Selection.IsActiveSpanReversed), (CInt(CLng(Fix(_textView.Selection.ActiveSpan.GetStartPoint(_textView.TextSnapshot))) Mod Integer.MaxValue)), (CInt(CLng(Fix(_textView.Selection.ActiveSpan.GetEndPoint(_textView.TextSnapshot))) Mod Integer.MaxValue))))
                        flag = caretPosition.TextInsertionIndex >= num
                    End If

                    caretPosition = (If((Not flag), _textView.Caret.MoveTo(extentOfCurrentWord.Start, CaretPlacement.LeadingEdgeOfCharacter), _textView.Caret.MoveTo(extentOfCurrentWord.Start + extentOfCurrentWord.Length, CaretPlacement.LeadingEdgeOfCharacter)))
                    Dim extentOfCurrentWord2 As TextExtent = GetExtentOfCurrentWord(num)
                    Dim activeSpan1 As ITextSpan

                    If flag Then
                        activeSpan1 = _textView.TextSnapshot.CreateTextSpan(Span.FromBounds(extentOfCurrentWord2.Start, caretPosition.TextInsertionIndex), SpanTrackingMode.EdgeExclusive)
                        _textView.Selection.IsActiveSpanReversed = False
                    Else
                        activeSpan1 = _textView.TextSnapshot.CreateTextSpan(Span.FromBounds(caretPosition.TextInsertionIndex, extentOfCurrentWord2.End), SpanTrackingMode.EdgeExclusive)
                        _textView.Selection.IsActiveSpanReversed = True
                    End If

                    _textView.Selection.ActiveSpan = activeSpan1
                    UpdateCaret()
                    _selectedWordSpan = _textView.TextSnapshot.CreateTextSpan(extentOfCurrentWord2.Start, extentOfCurrentWord2.Length, SpanTrackingMode.EdgeExclusive)
                    _selectedWordSpanIsReversed = _textView.Selection.IsActiveSpanReversed

                Else
                    UpdateCaret()
                    _textView.Selection.Clear()
                    SelectWordUnderCaret()
                End If

                result = True
            End If

            Return result
        End Function
    End Class
End Namespace
