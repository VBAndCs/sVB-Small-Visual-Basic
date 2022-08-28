Imports System.Globalization
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Windows
Imports Microsoft.Nautilus.Core.Undo
Imports Microsoft.Nautilus.Text.Diagnostics
Imports Microsoft.Nautilus.Text.Editor

Namespace Microsoft.Nautilus.Text.Operations
    Friend Class EditorOperations
        Implements IEditorOperations

        Private Enum CaretMovementDirection
            Previous
            [Next]
        End Enum

        Private Enum LetterCase
            Uppercase
            Lowercase
        End Enum

        Private _textStructureNavigator As ITextStructureNavigator

        Public ReadOnly Property CanPaste As Boolean Implements IEditorOperations.CanPaste
            Get
                Try
                    Return Clipboard.ContainsText()
                Catch __unusedExternalException1__ As ExternalException
                    Return False
                End Try
            End Get
        End Property

        Public ReadOnly Property TextView As ITextView Implements IEditorOperations.TextView

        Public Property ConvertTabsToSpace As Boolean Implements IEditorOperations.ConvertTabsToSpace

        Public Property OverwriteMode As Boolean Implements IEditorOperations.OverwriteMode

        Public Property TabSize As Integer Implements IEditorOperations.TabSize

        Friend Sub New(
                      textView As ITextView,
                      textStructureNavigatorFactory As ITextStructureNavigatorFactory)

            If textView Is Nothing Then
                Throw New ArgumentNullException("textView")
            End If

            If textStructureNavigatorFactory Is Nothing Then
                Throw New ArgumentNullException("textStructureNavigatorFactory")
            End If

            _TextView = textView
            _textStructureNavigator = textStructureNavigatorFactory.CreateTextStructureNavigator(_TextView.TextBuffer)
        End Sub

        Public Sub MoveCharacterRight([select] As Boolean) Implements IEditorOperations.MoveCharacterRight
            MoveCharacter(CaretMovementDirection.Next, [select])
        End Sub

        Public Sub MoveCharacterLeft([select] As Boolean) Implements IEditorOperations.MoveCharacterLeft
            MoveCharacter(CaretMovementDirection.Previous, [select])
        End Sub

        Public Sub MoveToNextWord([select] As Boolean) Implements IEditorOperations.MoveToNextWord
            Dim nextStartOfWord As Integer = GetNextStartOfWord(_textView.Caret.Position.TextInsertionIndex)
            If [select] Then
                ExtendSelection(nextStartOfWord)
            Else
                _textView.Selection.Clear()
            End If

            _textView.Caret.MoveTo(nextStartOfWord)
            _textView.Caret.EnsureVisible()
            _textView.Caret.CapturePreferredBounds()
        End Sub

        Public Sub MoveToPreviousWord([select] As Boolean) Implements IEditorOperations.MoveToPreviousWord
            Dim previousStartOfWord As Integer = GetPreviousStartOfWord(_textView.Caret.Position.TextInsertionIndex)
            If [select] Then
                ExtendSelection(previousStartOfWord)
            Else
                _textView.Selection.Clear()
            End If

            _textView.Caret.MoveTo(previousStartOfWord)
            _textView.Caret.EnsureVisible()
            _textView.Caret.CapturePreferredBounds()
        End Sub

        Public Sub MoveToStartOfDocument([select] As Boolean) Implements IEditorOperations.MoveToStartOfDocument
            Dim num As Integer = 0
            If [select] Then
                ExtendSelection(num)
            Else
                _textView.Selection.Clear()
            End If

            _textView.Caret.MoveTo(num)
            _textView.Caret.EnsureVisible()
            _textView.Caret.CapturePreferredBounds()
        End Sub

        Public Sub MoveToEndOfDocument([select] As Boolean) Implements IEditorOperations.MoveToEndOfDocument
            Dim length1 As Integer = _textView.TextSnapshot.Length
            If [select] Then
                ExtendSelection(length1)
            Else
                _textView.Selection.Clear()
            End If

            _textView.Caret.MoveTo(length1)
            _textView.Caret.EnsureVisible()
            _textView.Caret.CapturePreferredBounds()
        End Sub

        Public Sub MoveCurrentLineToTop() Implements IEditorOperations.MoveCurrentLineToTop
            Dim textInsertionIndex1 As Integer = _textView.Caret.Position.TextInsertionIndex
            _textView.DisplayTextLineContainingCharacter(textInsertionIndex1, 0.0, ViewRelativePosition.Top)
        End Sub

        Public Sub MoveCurrentLineToBottom() Implements IEditorOperations.MoveCurrentLineToBottom
            Dim textInsertionIndex1 As Integer = _textView.Caret.Position.TextInsertionIndex
            _textView.DisplayTextLineContainingCharacter(textInsertionIndex1, 0.0, ViewRelativePosition.Bottom)
        End Sub

        Public Sub DeleteWordToRight(undoHistory1 As UndoHistory) Implements IEditorOperations.DeleteWordToRight
            Dim textInsertionIndex1 As Integer = _textView.Caret.Position.TextInsertionIndex
            Dim length1 As Integer = GetNextStartOfWord(textInsertionIndex1) - textInsertionIndex1
            Delete(New Span(textInsertionIndex1, length1), undoHistory1, "Delete Word To Right", TextEditAction.Delete)
        End Sub

        Public Sub DeleteWordToLeft(undoHistory1 As UndoHistory) Implements IEditorOperations.DeleteWordToLeft
            Dim insertionIndex As Integer = _TextView.Caret.Position.TextInsertionIndex
            Dim startOfWord As Integer = GetPreviousStartOfWord(insertionIndex)
            Dim s As New Span(startOfWord, insertionIndex - startOfWord)
            Delete(s, undoHistory1, "Delete Word To Left", TextEditAction.Delete)
        End Sub

        Public Sub YankCurrentLine(undoHistory1 As UndoHistory) Implements IEditorOperations.YankCurrentLine
            Dim start = 0
            Dim [end] = 0

            If _TextView.Selection.IsEmpty Then
                Dim line = _TextView.TextSnapshot.GetLineFromPosition(_TextView.Caret.Position.TextInsertionIndex)
                start = line.Start
                [end] = line.EndIncludingLineBreak

            Else
                Dim selStart = _TextView.Selection.ActiveSnapshotSpan.Start
                Dim selEnd = _TextView.Selection.ActiveSnapshotSpan.End

                If _TextView.Selection.IsActiveSpanReversed Then
                    Dim temp = selEnd
                    selEnd = selStart
                    selStart = temp
                End If

                start = _TextView.TextSnapshot.GetLineFromPosition(selStart).Start
                [end] = _TextView.TextSnapshot.GetLineFromPosition(selEnd).EndIncludingLineBreak
            End If

            _TextView.Selection.ActiveSnapshotSpan = New SnapshotSpan(_TextView.TextSnapshot, start, [end] - start)
            CutSelection(undoHistory1)
        End Sub

        Public Sub DeleteCharacterToLeft(undoHistory1 As UndoHistory) Implements IEditorOperations.DeleteCharacterToLeft
            Dim textInsertionIndex1 As Integer = _textView.Caret.Position.TextInsertionIndex
            Dim previousCodePoint As Integer = GetPreviousCodePoint(textInsertionIndex1)

            If _TextView.Caret.Position.Placement = CaretPlacement.TrailingEdgeOfCharacter Then
                _TextView.Caret.MoveTo(_TextView.Caret.Position.TextInsertionIndex, CaretPlacement.LeadingEdgeOfCharacter)
            End If

            Delete(New Span(previousCodePoint, textInsertionIndex1 - previousCodePoint), undoHistory1, "Delete Character To Left", TextEditAction.Backspace)
        End Sub

        Public Sub DeleteCharacterToRight(undoHistory As UndoHistory) Implements IEditorOperations.DeleteCharacterToRight
            Dim insertionIndex As Integer = _TextView.Caret.Position.TextInsertionIndex
            Dim length As Integer = GetNextCodePoint(insertionIndex) - insertionIndex
            Dim s As New Span(insertionIndex, length)
            Delete(s, undoHistory, "Delete Character To Right", TextEditAction.Backspace)
        End Sub

        Private Sub Delete(span1 As Span, undoHistory As UndoHistory, undoTransactionText As String, action As TextEditAction)
            EditorTrace.TraceTextDeleteStart()
            If undoHistory Is Nothing Then
                Throw New ArgumentNullException("undoHistory")
            End If

            Using undoTransaction1 As UndoTransaction = undoHistory.CreateTransaction(undoTransactionText)
                Dim flag As Boolean = DeleteSelection(undoHistory)

                If Not flag AndAlso span1.Length <> 0 Then
                    flag = ReplaceText(span1, String.Empty, action, undoHistory)
                End If

                If flag Then
                    undoTransaction1.Complete()
                Else
                    undoTransaction1.Cancel()
                End If

            End Using

            _TextView.Caret.EnsureVisible()
            _TextView.Caret.CapturePreferredBounds()
            EditorTrace.TraceTextDeleteEnd()
        End Sub

        Public Sub SelectCurrentWord() Implements IEditorOperations.SelectCurrentWord
            EditorTrace.TraceTextSelectStart()
            Dim span = Me.GetCurrentWordSpan()
            _TextView.Selection.ActiveSnapshotSpan = New SnapshotSpan(_TextView.TextSnapshot, span)
            _TextView.Caret.MoveTo(span.End)
            _TextView.Caret.EnsureVisible()
            _TextView.Caret.CapturePreferredBounds()

            EditorTrace.TraceTextSelectEnd()
        End Sub

        Private Function GetCurrentWordSpan() As Span Implements IEditorOperations.GetCurrentWordSpan
            Dim pos As New SnapshotPoint(_TextView.TextSnapshot,
                                If(_TextView.Selection.IsEmpty,
                                       _TextView.Caret.Position.TextInsertionIndex,
                                       _TextView.Selection.ActiveSnapshotSpan.Start
                                )
                           )

            Return GetExtentOfCurrentWord(pos).Span
        End Function


        Private Function GetCurrentWord() As String Implements IEditorOperations.GetCurrentWord
            Return TextView.TextSnapshot.GetText(GetCurrentWordSpan())
        End Function

        Public Sub SelectAll() Implements IEditorOperations.SelectAll
            EditorTrace.TraceTextSelectStart()

            Dim length As Integer = _TextView.TextSnapshot.Length
            _TextView.Selection.ActiveSnapshotSpan = New SnapshotSpan(_TextView.TextSnapshot, New Span(0, length))
            _TextView.Caret.MoveTo(length)
            _TextView.Caret.EnsureVisible()

            EditorTrace.TraceTextSelectEnd()
        End Sub

        Public Sub ExtendSelection(newEnd As Integer) Implements IEditorOperations.ExtendSelection
            EditorTrace.TraceTextSelectStart()

            If newEnd < 0 Then
                Throw New ArgumentOutOfRangeException("newEnd")
            End If

            Dim selection1 As ITextSelection = _TextView.Selection
            Dim num As Integer = (If(selection1.IsEmpty, _TextView.Caret.Position.TextInsertionIndex, (If((Not selection1.IsActiveSpanReversed), selection1.ActiveSnapshotSpan.Start, selection1.ActiveSnapshotSpan.End))))

            If num < newEnd Then
                selection1.ActiveSnapshotSpan = New SnapshotSpan(_TextView.TextSnapshot, Span.FromBounds(num, newEnd))
                selection1.IsActiveSpanReversed = False
            Else
                selection1.ActiveSnapshotSpan = New SnapshotSpan(_textView.TextSnapshot, Span.FromBounds(newEnd, num))
                selection1.IsActiveSpanReversed = True
            End If

            EditorTrace.TraceTextSelectEnd()
        End Sub

        Public Sub MoveCaretAndExtendSelection(textLine As ITextLine, horizontalOffset As Double) Implements IEditorOperations.MoveCaretAndExtendSelection
            EditorTrace.TraceTextSelectStart()

            If textLine Is Nothing Then
                Throw New ArgumentNullException("textLine")
            End If

            Dim selection1 As ITextSelection = _textView.Selection
            Dim num As Integer = (If(selection1.IsEmpty, _textView.Caret.Position.TextInsertionIndex, (If((Not selection1.IsActiveSpanReversed), selection1.ActiveSnapshotSpan.Start, selection1.ActiveSnapshotSpan.End))))
            Dim textInsertionIndex1 As Integer = textLine.MoveCaretToLocation(horizontalOffset).TextInsertionIndex

            If num < textInsertionIndex1 Then
                selection1.ActiveSnapshotSpan = New SnapshotSpan(_TextView.TextSnapshot, Span.FromBounds(num, textInsertionIndex1))
                selection1.IsActiveSpanReversed = False
            Else
                selection1.ActiveSnapshotSpan = New SnapshotSpan(_textView.TextSnapshot, Span.FromBounds(textInsertionIndex1, num))
                selection1.IsActiveSpanReversed = True
            End If

            EditorTrace.TraceTextSelectEnd()
        End Sub

        Public Sub MoveLineUp([select] As Boolean) Implements IEditorOperations.MoveLineUp
            _textView.Caret.EnsureVisible()

            Dim formattedTextLines1 As IFormattedTextLineCollection = _TextView.FormattedTextLines
            Dim textLineContainingPosition As ITextLine = formattedTextLines1.GetTextLineContainingPosition(_textView.Caret.Position.CharacterIndex)
            Dim num As Integer = formattedTextLines1.IndexOf(textLineContainingPosition)

            If num <> 0 Then
                If [select] Then
                    MoveCaretAndExtendSelection(formattedTextLines1(num - 1), _TextView.Caret.PreferredBounds.Left)
                Else
                    If _TextView.Selection.IsEmpty Then
                        formattedTextLines1(num - 1).MoveCaretToLocation(_TextView.Caret.PreferredBounds.Left)
                    Else
                        Dim textLineContainingPosition2 As ITextLine = formattedTextLines1.GetTextLineContainingPosition(_TextView.Selection.ActiveSnapshotSpan.Start)
                        Dim num2 As Integer = formattedTextLines1.IndexOf(textLineContainingPosition2)
                        If num2 > 0 Then
                            formattedTextLines1(num2 - 1).MoveCaretToLocation(_TextView.Caret.PreferredBounds.Left)
                        Else
                            _TextView.Caret.MoveTo(_TextView.Selection.ActiveSnapshotSpan.Start)
                        End If
                    End If
                    _TextView.Selection.Clear()
                End If
                _TextView.Caret.EnsureVisible()
            End If

            _TextView.Caret.CapturePreferredVerticalBounds()
        End Sub

        Public Sub MoveLineDown([select] As Boolean) Implements IEditorOperations.MoveLineDown
            _textView.Caret.EnsureVisible()

            Dim formattedTextLines1 = _TextView.FormattedTextLines
            Dim textLineContainingPosition = formattedTextLines1.GetTextLineContainingPosition(_TextView.Caret.Position.CharacterIndex)
            Dim num As Integer = formattedTextLines1.IndexOf(textLineContainingPosition)

            If num <> formattedTextLines1.Count - 1 Then
                If [select] Then
                    MoveCaretAndExtendSelection(formattedTextLines1(num + 1), _TextView.Caret.PreferredBounds.Left)
                Else
                    If _TextView.Selection.IsEmpty Then
                        formattedTextLines1(num + 1).MoveCaretToLocation(_TextView.Caret.PreferredBounds.Left)
                    Else
                        Dim textLineContainingPosition2 As ITextLine = formattedTextLines1.GetTextLineContainingPosition(_TextView.Selection.ActiveSnapshotSpan.End)
                        Dim num2 As Integer = formattedTextLines1.IndexOf(textLineContainingPosition2)
                        If num2 <> formattedTextLines1.Count - 1 Then
                            formattedTextLines1(num2 + 1).MoveCaretToLocation(_TextView.Caret.PreferredBounds.Left)
                        Else
                            formattedTextLines1(num2).MoveCaretToLocation(textLineContainingPosition2.Right)
                        End If
                    End If
                    _TextView.Selection.Clear()
                End If
                _TextView.Caret.EnsureVisible()
            End If

            _TextView.Caret.CapturePreferredVerticalBounds()
        End Sub

        Public Sub PageUp([select] As Boolean) Implements IEditorOperations.PageUp
            EditorTrace.TraceTextPageUpStart()
            PageUpDown([select], ScrollDirection.Up)
            EditorTrace.TraceTextPageUpEnd()
        End Sub

        Public Sub PageDown([select] As Boolean) Implements IEditorOperations.PageDown
            EditorTrace.TraceTextPageDownStart()
            PageUpDown([select], ScrollDirection.Down)
            EditorTrace.TraceTextPageDownEnd()
        End Sub

        Private Sub PageUpDown([select] As Boolean, direction As ScrollDirection)
            Dim flag As Boolean = False
            Dim num As Double = _textView.Caret.PreferredBounds.Top + _textView.Caret.PreferredBounds.Height * 0.5

            If direction = ScrollDirection.Up Then
                Dim formattedTextLines1 As IFormattedTextLineCollection = _TextView.FormattedTextLines
                Dim firstFullyVisibleLine1 As ITextLine = formattedTextLines1.FirstFullyVisibleLine
                _TextView.ViewScroller.ScrollViewportVerticallyByPage(ScrollDirection.Up)
                Dim num2 As Double = _TextView.ViewportHeight - firstFullyVisibleLine1.Bottom
                If num2 > 0.0 Then
                    flag = True
                    num -= num2
                End If
            Else
                Dim formattedTextLines2 As IFormattedTextLineCollection = _textView.FormattedTextLines
                Dim lastFullyVisibleLine1 As ITextLine = formattedTextLines2.LastFullyVisibleLine
                _textView.ViewScroller.ScrollViewportVerticallyByPage(ScrollDirection.Down)
                Dim bottom1 As Double = lastFullyVisibleLine1.Bottom
                If bottom1 > 0.0 Then
                    flag = True
                    num += bottom1
                End If
            End If

            Dim formattedTextLines3 As IFormattedTextLineCollection = _textView.FormattedTextLines
            Dim index As Integer

            If Not formattedTextLines3.FindTextLineIndexContainingYCoordinate(num, index) Then
                If index < 0 Then index = 0
                flag = True
            End If

            Dim textLine As ITextLine = formattedTextLines3(index)
            If formattedTextLines3.GetVisibilityStateOfTextLine(textLine) <> VisibilityState.FullyVisible Then
                If index > 0 AndAlso formattedTextLines3.GetVisibilityStateOfTextLine(formattedTextLines3(index - 1)) = VisibilityState.FullyVisible Then
                    textLine = formattedTextLines3(index - 1)
                ElseIf index < formattedTextLines3.Count - 1 AndAlso formattedTextLines3.GetVisibilityStateOfTextLine(formattedTextLines3(index + 1)) = VisibilityState.FullyVisible Then
                    textLine = formattedTextLines3(index + 1)
                End If
            End If

            If [select] Then
                MoveCaretAndExtendSelection(textLine, _textView.Caret.PreferredBounds.Left)
            Else
                textLine.MoveCaretToLocation(_textView.Caret.PreferredBounds.Left)
                _textView.Selection.Clear()
            End If

            _textView.Caret.EnsureVisible()
            If flag Then
                _textView.Caret.CapturePreferredVerticalBounds()
            End If
        End Sub

        Public Sub MoveToEndOfLine([select] As Boolean) Implements IEditorOperations.MoveToEndOfLine
            Dim positionOfEndOfLine As Integer = _textStructureNavigator.GetPositionOfEndOfLine(New SnapshotPoint(_textView.TextSnapshot, _textView.Caret.Position.TextInsertionIndex))
            If [select] Then
                ExtendSelection(positionOfEndOfLine)
            Else
                _textView.Selection.Clear()
            End If
            _textView.Caret.MoveTo(positionOfEndOfLine)
            _textView.Caret.EnsureVisible()
            _textView.Caret.CapturePreferredBounds()
        End Sub

        Public Sub MoveToStartOfLine([select] As Boolean) Implements IEditorOperations.MoveToStartOfLine
            Dim positionOfStartOfLine As Integer = _textStructureNavigator.GetPositionOfStartOfLine(New SnapshotPoint(_textView.TextSnapshot, _textView.Caret.Position.TextInsertionIndex))
            If [select] Then
                ExtendSelection(positionOfStartOfLine)
            Else
                _textView.Selection.Clear()
            End If
            _textView.Caret.MoveTo(positionOfStartOfLine)
            _textView.Caret.EnsureVisible()
            _textView.Caret.CapturePreferredBounds()
        End Sub

        Public Sub InsertNewline(undoHistory1 As UndoHistory) Implements IEditorOperations.InsertNewline
            If undoHistory1 Is Nothing Then
                Throw New ArgumentNullException("undoHistory")
            End If

            Dim sb As New StringBuilder(Microsoft.VisualBasic.vbNewLine)
            Dim insertionIndex = _TextView.Caret.Position.TextInsertionIndex
            Dim line = _TextView.TextSnapshot.GetLineFromPosition(insertionIndex)

            If insertionIndex <> line.Start Then
                Dim text As String = line.GetText()
                For i As Integer = 0 To text.Length - 1
                    If text(i) = vbTab Then
                        sb.Append(vbTab)
                        Continue For
                    End If
                    If text(i) <> " "c Then
                        Exit For
                    End If
                    sb.Append(" "c)
                Next
            End If

            Using undoTransaction = undoHistory1.CreateTransaction("Insert Newline")
                If ReplaceSelection(sb.ToString(), TextEditAction.Enter, undoHistory1) Then
                    undoTransaction.Complete()
                Else
                    undoTransaction.Cancel()
                End If
            End Using

            _TextView.Caret.EnsureVisible()
        End Sub

        Public Sub InsertTab(undoHistory As UndoHistory) Implements IEditorOperations.InsertTab
            If undoHistory Is Nothing Then
                Throw New ArgumentNullException("undoHistory")
            End If

            Using undoTransaction = undoHistory.CreateTransaction("Insert Tab")
                Dim text As String

                If ConvertTabsToSpace Then
                    Dim columnOffset = ColumnOffsetOfPositionInLine(_TextView.Caret.Position.TextInsertionIndex, _TextView.TextSnapshot.GetLineFromPosition(_TextView.Caret.Position.TextInsertionIndex))
                    text = New String(" "c, TabSize - columnOffset Mod TabSize)
                Else
                    text = vbTab
                End If

                Dim done = If(_TextView.Selection.IsEmpty AndAlso _OverwriteMode,
                         InsertTextOverwriteMode(_TextView.Caret.Position.TextInsertionIndex, text, undoHistory),
                         ReplaceSelection(text, TextEditAction.Type, undoHistory))

                If done Then
                    undoTransaction.Complete()
                Else
                    undoTransaction.Cancel()
                End If
            End Using

        End Sub

        Public Sub RemovePreviousTab(undoHistory1 As UndoHistory) Implements IEditorOperations.RemovePreviousTab
            If undoHistory1 Is Nothing Then
                Throw New ArgumentNullException("undoHistory")
            End If

            Dim text As String = ""
            Dim textInsertionIndex1 As Integer = _textView.Caret.Position.TextInsertionIndex

            If textInsertionIndex1 = 0 Then Return

            Dim num As Integer = ColumnOffsetOfPositionInLine(textInsertionIndex1, _textView.TextSnapshot.GetLineFromPosition(textInsertionIndex1))
            If ConvertTabsToSpace Then
                Dim num2 As Integer = num Mod TabSize

                If num2 = 0 Then num2 = TabSize

                Dim num3 As Integer = 0
                Dim num4 As Integer = textInsertionIndex1 - 1

                While num4 >= textInsertionIndex1 - num2 AndAlso " "c = _TextView.TextSnapshot(num4)
                    num3 += 1
                    num4 -= 1
                End While

                If num3 > 0 Then
                    text = New String(" "c, num3)
                End If

            Else
                Dim text2 As String = vbTab
                Dim text3 As String = _textView.TextSnapshot.GetText(_textView.TextSnapshot.GetLineFromPosition(textInsertionIndex1).Start, num)
                Dim num5 As Integer = text3.LastIndexOf(text2, num, StringComparison.CurrentCulture)

                If num5 <> -1 AndAlso num5 + text2.Length = num Then
                    text = text2
                End If
            End If

            If text.Length <= 0 Then Return

            Using undoTransaction1 As UndoTransaction = undoHistory1.CreateTransaction("Remove Previous Tab")
                Dim num6 As Integer = textInsertionIndex1 - text.Length

                Using textEdit As ITextEdit = _TextView.TextBuffer.CreateEdit()
                    If textEdit.CanDeleteOrReplace(New Span(num6, text.Length)) Then
                        Dim undo As New BeforeTextBufferChangeUndoPrimitive(_TextView, _TextView.Caret.Position.CharacterIndex, _TextView.Caret.Position.Placement, _TextView.Selection.ActiveSnapshotSpan)
                        undoHistory1.CurrentTransaction.AddUndo(undo)
                        textEdit.Replace(num6, text.Length, "")
                        textEdit.Apply()
                        Dim undo2 As New AfterTextBufferChangeUndoPrimitive(_TextView, _TextView.Caret.Position.CharacterIndex, _TextView.Caret.Position.Placement, _TextView.Selection.ActiveSnapshotSpan)
                        undoHistory1.CurrentTransaction.AddUndo(undo2)
                        undoTransaction1.Complete()
                    Else
                        undoTransaction1.Cancel()
                    End If
                End Using

            End Using
        End Sub

        Public Sub IndentSelection(undoHistory1 As UndoHistory) Implements IEditorOperations.IndentSelection
            If undoHistory1 Is Nothing Then
                Throw New ArgumentNullException("undoHistory")
            End If

            If _textView.Selection.IsEmpty Then
                Throw New InvalidOperationException
            End If

            Dim textSnapshot1 As ITextSnapshot = _textView.TextSnapshot
            Dim span1 As Span = _textView.Selection.ActiveSnapshotSpan
            Dim line1 = textSnapshot1.GetLineFromPosition(span1.Start)
            Dim line2 = textSnapshot1.GetLineFromPosition(span1.End)
            Dim num As Integer = line2.LineNumber

            If line1.End >= span1.End Then
                InsertTab(undoHistory1)
                Return
            End If

            Dim text As String = (If(ConvertTabsToSpace, New String(" "c, TabSize), vbTab))
            If span1.End = line2.Start Then
                num -= 1
            End If

            Using undoTransaction1 As UndoTransaction = undoHistory1.CreateTransaction("Indent Selection")
                Dim span2 As Span = span1
                Dim position1 As ICaretPosition = _TextView.Caret.Position
                Dim undo As New BeforeTextBufferChangeUndoPrimitive(_TextView, _TextView.Caret.Position.CharacterIndex, _TextView.Caret.Position.Placement, span1)
                undoHistory1.CurrentTransaction.AddUndo(undo)

                Dim flag As Boolean = False
                Using textEdit As ITextEdit = _TextView.TextBuffer.CreateEdit()
                    For i As Integer = line1.LineNumber To num
                        Dim lineFromLineNumber As ITextSnapshotLine = textSnapshot1.GetLineFromLineNumber(i)
                        If lineFromLineNumber.Length <> 0 Then
                            Dim start1 As Integer = lineFromLineNumber.Start
                            If Not textEdit.CanInsert(start1) Then
                                flag = True
                                Exit For
                            End If
                            textEdit.Insert(start1, text)
                        End If
                    Next

                    If flag Then
                        undoTransaction1.Cancel()
                        Return
                    End If

                    textEdit.Apply()
                    Dim snapshotSpan1 As SnapshotSpan = _TextView.Selection.ActiveSnapshotSpan
                    Dim num2 As Integer = _TextView.Caret.Position.CharacterIndex
                    Dim caretPlacement1 As CaretPlacement = _TextView.Caret.Position.Placement
                    Dim positionOfNextNonWhiteSpaceCharacter As Integer = line1.GetPositionOfNextNonWhiteSpaceCharacter(0)
                    Dim positionOfNextNonWhiteSpaceCharacter2 As Integer = line2.GetPositionOfNextNonWhiteSpaceCharacter(0)
                    Dim flag2 As Boolean = span2.Start <= line1.Start + positionOfNextNonWhiteSpaceCharacter
                    Dim flag3 As Boolean = span2.End < line2.Start + positionOfNextNonWhiteSpaceCharacter2

                    If flag2 OrElse flag3 Then
                        Dim num3 As Integer = (If(flag2, span2.Start, _TextView.Selection.ActiveSnapshotSpan.Start))
                        Dim num4 As Integer = (If((Not flag3 OrElse span2.End = line2.Start OrElse line2.Length = 0), _TextView.Selection.ActiveSnapshotSpan.End, (_TextView.Selection.ActiveSnapshotSpan.End - text.Length)))
                        snapshotSpan1 = New SnapshotSpan(_TextView.TextSnapshot, num3, num4 - num3)
                        If position1.TextInsertionIndex = span2.Start Then
                            num2 = num3
                            caretPlacement1 = CaretPlacement.LeadingEdgeOfCharacter
                        Else
                            num2 = num4
                            caretPlacement1 = CaretPlacement.LeadingEdgeOfCharacter
                        End If
                        _TextView.Selection.ActiveSnapshotSpan = snapshotSpan1
                        _TextView.Caret.MoveTo(num2, caretPlacement1)
                    End If

                    _TextView.Caret.EnsureVisible()
                    _TextView.Caret.CapturePreferredBounds()
                    Dim undo2 As New AfterTextBufferChangeUndoPrimitive(_textView, num2, caretPlacement1, snapshotSpan1)
                    undoHistory1.CurrentTransaction.AddUndo(undo2)
                    undoTransaction1.Complete()
                End Using

            End Using

        End Sub

        Public Sub UnindentSelection(undoHistory1 As UndoHistory) Implements IEditorOperations.UnindentSelection
            If undoHistory1 Is Nothing Then
                Throw New ArgumentNullException("undoHistory")
            End If

            If _textView.Selection.IsEmpty Then
                Throw New InvalidOperationException
            End If

            Dim textSnapshot1 = _TextView.TextSnapshot
            Dim oldSelectionSpan As Span = _textView.Selection.ActiveSnapshotSpan
            Dim lineFromPosition = textSnapshot1.GetLineFromPosition(oldSelectionSpan.Start)
            Dim lineFromPosition2 = textSnapshot1.GetLineFromPosition(oldSelectionSpan.End)
            Dim num = lineFromPosition2.LineNumber
            If lineFromPosition.End >= oldSelectionSpan.End Then
                Return
            End If

            Dim text As String = (If(ConvertTabsToSpace, New String(" "c, TabSize), vbTab))
            If oldSelectionSpan.End = lineFromPosition2.Start Then
                num -= 1
            End If

            Dim num2 As Integer = (If((Not ConvertTabsToSpace), 1, TabSize))
            Using undoTransaction1 As UndoTransaction = undoHistory1.CreateTransaction("Unindent Selection")
                Dim undo As New BeforeTextBufferChangeUndoPrimitive(_textView, _textView.Caret.Position.CharacterIndex, _textView.Caret.Position.Placement, oldSelectionSpan)
                undoHistory1.CurrentTransaction.AddUndo(undo)
                Dim flag As Boolean = False

                Using textEdit As ITextEdit = _TextView.TextBuffer.CreateEdit()
                    For i As Integer = lineFromPosition.LineNumber To num
                        Dim start1 As Integer = textSnapshot1.GetLineFromLineNumber(i).Start
                        If textSnapshot1.Length <= start1 + num2 Then
                            Continue For
                        End If

                        Dim text2 As String = textSnapshot1.GetText(start1, num2)
                        If text2.IndexOf(text, StringComparison.CurrentCulture) = 0 Then
                            If Not textEdit.CanDeleteOrReplace(New Span(start1, text.Length)) Then
                                flag = True
                                Exit For
                            End If
                            textEdit.Replace(start1, text.Length, "")
                        End If
                    Next

                    If flag Then
                        undoTransaction1.Cancel()
                        Return
                    End If

                    textEdit.Apply()
                    Dim undo2 As New AfterTextBufferChangeUndoPrimitive(_TextView, _TextView.Caret.Position.CharacterIndex, _TextView.Caret.Position.Placement, _TextView.Selection.ActiveSnapshotSpan)
                    undoHistory1.CurrentTransaction.AddUndo(undo2)
                    undoTransaction1.Complete()
                End Using
            End Using
        End Sub

        Public Sub InsertText(text As String, undoHistory As UndoHistory) Implements IEditorOperations.InsertText
            If undoHistory Is Nothing Then
                Throw New ArgumentNullException(NameOf(undoHistory))
            End If

            EditorTrace.TraceTextInsertStart()
            If text Is Nothing Then
                Throw New ArgumentNullException("text")
            End If

            If undoHistory Is Nothing Then
                Throw New ArgumentNullException("undoHistory")
            End If

            If Not _TextView.Selection.IsEmpty Then
                Using undoTransaction = undoHistory.CreateTransaction("Stop Text Merge", isHidden:=True)
                    undoTransaction.Complete()
                End Using
            End If

            Using undoTransaction = undoHistory.CreateTransaction("Insert Text")
                undoTransaction.MergePolicy = New TextTransactionMergePolicy
                Dim done = If((_TextView.Selection.IsEmpty AndAlso _OverwriteMode),
                                        InsertTextOverwriteMode(_TextView.Caret.Position.TextInsertionIndex, text, undoHistory),
                                        ReplaceSelection(text, TextEditAction.Type, undoHistory)
                                    )
                If done Then
                    undoTransaction.Complete()
                Else
                    undoTransaction.Cancel()
                End If
            End Using

            EditorTrace.TraceTextInsertEnd()
        End Sub

        Public Sub SelectLine(lineNumber As Integer, extendSelection As Boolean) Implements IEditorOperations.SelectLine
            Dim snapshot = _TextView.TextSnapshot
            If lineNumber < 0 OrElse lineNumber > snapshot.LineCount Then
                Throw New ArgumentOutOfRangeException("lineNumber")
            End If

            Dim line = snapshot.GetLineFromLineNumber(lineNumber)
            If Not extendSelection OrElse _TextView.Selection.IsEmpty Then
                _TextView.Selection.ActiveSnapshotSpan = New SnapshotSpan(snapshot, New Span(line.Start, line.LengthIncludingLineBreak))
                _TextView.Selection.IsActiveSpanReversed = False
                _TextView.Caret.MoveTo(line.EndIncludingLineBreak)
                _TextView.Caret.EnsureVisible()

            Else
                Dim span = _TextView.Selection.ActiveSnapshotSpan
                Dim start = span.Start
                Dim [end] = span.End

                If _TextView.Selection.IsActiveSpanReversed Then
                    start = Math.Min(snapshot.Length, snapshot.GetLineFromPosition([end]).EndIncludingLineBreak + 1)
                Else
                    [end] = snapshot.GetLineFromPosition(start).EndIncludingLineBreak
                End If

                If line.EndIncludingLineBreak < [end] Then
                    _TextView.Selection.ActiveSnapshotSpan = New SnapshotSpan(snapshot, Text.Span.FromBounds(line.Start, [end]))
                    _TextView.Selection.IsActiveSpanReversed = True
                    _TextView.Caret.MoveTo(line.Start)
                    _TextView.Caret.EnsureVisible()
                Else
                    Dim start2 = Math.Min(Math.Min(start, [end]), line.EndIncludingLineBreak)
                    Dim end2 = Math.Max(Math.Max(start, [end]), line.EndIncludingLineBreak)
                    _TextView.Selection.ActiveSnapshotSpan = New SnapshotSpan(snapshot, Text.Span.FromBounds(start2, end2))
                    _TextView.Selection.IsActiveSpanReversed = False
                    _TextView.Caret.MoveTo(end2)
                    _TextView.Caret.EnsureVisible()
                End If
            End If

            _TextView.Caret.CapturePreferredBounds()
        End Sub

        Public Sub ResetSelection() Implements IEditorOperations.ResetSelection
            _textView.Selection.Clear()
        End Sub

        Public Function DeleteSelection(undoHistory1 As UndoHistory) As Boolean Implements IEditorOperations.DeleteSelection
            If undoHistory1 Is Nothing Then
                Throw New ArgumentNullException("undoHistory")
            End If

            If Not _textView.Selection.IsEmpty Then
                Using undoTransaction1 As UndoTransaction = undoHistory1.CreateTransaction("Delete Text")
                    ReplaceSelection(String.Empty, TextEditAction.Delete, undoHistory1)
                    undoTransaction1.Complete()
                End Using
                Return True
            End If

            Return False
        End Function

        Public Sub ReplaceSelection(text As String, undoHistory1 As UndoHistory) Implements IEditorOperations.ReplaceSelection
            If text Is Nothing Then
                Throw New ArgumentNullException("text")
            End If

            If undoHistory1 Is Nothing Then
                Throw New ArgumentNullException("undoHistory")
            End If

            Using undoTransaction1 As UndoTransaction = undoHistory1.CreateTransaction("Replace Selection With " & text)
                If ReplaceSelection(text, TextEditAction.Type, undoHistory1) Then
                    undoTransaction1.Complete()
                Else
                    undoTransaction1.Cancel()
                End If
            End Using
        End Sub

        Public Sub ReplaceText(span1 As Span, text As String, undoHistory1 As UndoHistory) Implements IEditorOperations.ReplaceText
            If span1.End > _textView.TextSnapshot.Length Then
                Throw New ArgumentOutOfRangeException("span")
            End If

            If undoHistory1 Is Nothing Then
                Throw New ArgumentNullException("undoHistory")
            End If

            Using undoTransaction1 As UndoTransaction = undoHistory1.CreateTransaction("Replace Text")
                If ReplaceText(span1, text, TextEditAction.Type, undoHistory1) Then
                    undoTransaction1.Complete()
                Else
                    undoTransaction1.Cancel()
                End If
            End Using
        End Sub

        Public Function ReplaceAllMatches(searchText As String, replaceText1 As String, matchCase As Boolean, matchWholeWord As Boolean, useRegularExpressions As Boolean, undoHistory1 As UndoHistory) As Integer Implements IEditorOperations.ReplaceAllMatches
            If searchText Is Nothing Then
                Throw New ArgumentNullException("searchText")
            End If

            If undoHistory1 Is Nothing Then
                Throw New ArgumentNullException("undoHistory")
            End If

            Return 0
        End Function

        Public Sub CopySelection() Implements IEditorOperations.CopySelection
            EditorTrace.TraceTextCopyStart()

            If Not _TextView.Selection.IsEmpty Then
                Dim avalonTextView = TryCast(_TextView, IAvalonTextView)
                If avalonTextView IsNot Nothing AndAlso EditorCommands.Current IsNot Nothing Then
                    EditorCommands.Current.CopyHandler(avalonTextView)
                Else
                    Dim text = _TextView.Selection.ActiveSnapshotSpan.GetText()
                    Try
                        Clipboard.SetDataObject(New DataObject(GetType(String), text), copy:=True)
                    Catch
                    End Try
                End If
            End If

            EditorTrace.TraceTextCopyEnd()
        End Sub

        Public Sub CutSelection(undoHistory1 As UndoHistory) Implements IEditorOperations.CutSelection
            EditorTrace.TraceTextCutStart()
            If undoHistory1 Is Nothing Then
                Throw New ArgumentNullException("undoHistory")
            End If

            CopySelection()

            Using undoTransaction1 As UndoTransaction = undoHistory1.CreateTransaction("Cut Selection")
                If DeleteSelection(undoHistory1) Then
                    _textView.Caret.EnsureVisible()
                    undoTransaction1.Complete()
                Else
                    undoTransaction1.Cancel()
                End If
            End Using

            EditorTrace.TraceTextCutEnd()
        End Sub

        Public Sub Paste(undoHistory1 As UndoHistory) Implements IEditorOperations.Paste
            EditorTrace.TraceTextPasteStart()
            If undoHistory1 Is Nothing Then
                Throw New ArgumentNullException("undoHistory")
            End If

            Dim text As String = Nothing
            Try
                Dim dataObject1 As IDataObject = Clipboard.GetDataObject()
                If dataObject1 Is Nothing OrElse Not dataObject1.GetDataPresent(GetType(String)) Then
                    Return
                End If

                text = CStr(dataObject1.GetData(DataFormats.UnicodeText))
                If text Is Nothing Then
                    text = CStr(dataObject1.GetData(DataFormats.Text))
                End If

            Catch

            End Try

            If text IsNot Nothing Then
                Using undoTransaction1 As UndoTransaction = undoHistory1.CreateTransaction("Paste")
                    If ReplaceSelection(text, TextEditAction.Paste, undoHistory1) Then
                        _textView.Caret.EnsureVisible()
                        _textView.Caret.CapturePreferredBounds()
                        undoTransaction1.Complete()
                    Else
                        undoTransaction1.Cancel()
                    End If
                End Using
            End If

            EditorTrace.TraceTextPasteEnd()
        End Sub

        Public Sub GotoLine(lineNumber1 As Integer) Implements IEditorOperations.GotoLine
            If lineNumber1 < 0 OrElse lineNumber1 > _textView.TextSnapshot.LineCount - 1 Then
                Throw New ArgumentOutOfRangeException("lineNumber")
            End If

            Dim start1 As Integer = _textView.TextSnapshot.GetLineFromLineNumber(lineNumber1).Start
            _textView.Caret.MoveTo(start1)
            _textView.Caret.EnsureVisible()
            _textView.Caret.CapturePreferredBounds()
        End Sub

        Public Sub ScrollUpAndMoveCaretIfNecessary() Implements IEditorOperations.ScrollUpAndMoveCaretIfNecessary
            Dim formattedTextLines1 As IFormattedTextLineCollection = _textView.FormattedTextLines
            _textView.ViewScroller.ScrollViewportVerticallyByLine(ScrollDirection.Up)

            If formattedTextLines1.FirstFullyVisibleLine.LineSpan.Start > _TextView.Caret.Position.TextInsertionIndex OrElse _TextView.Caret.Position.TextInsertionIndex > formattedTextLines1.LastFullyVisibleLine.LineSpan.End Then
                _TextView.Selection.Clear()
                formattedTextLines1.LastFullyVisibleLine.MoveCaretToLocation(_TextView.Caret.PreferredBounds.Left)
                _TextView.Caret.EnsureVisible()
            End If
        End Sub

        Public Sub ScrollDownAndMoveCaretIfNecessary() Implements IEditorOperations.ScrollDownAndMoveCaretIfNecessary
            Dim formattedTextLines1 As IFormattedTextLineCollection = _textView.FormattedTextLines
            _textView.ViewScroller.ScrollViewportVerticallyByLine(ScrollDirection.Down)

            If formattedTextLines1.FirstFullyVisibleLine.LineSpan.Start > _TextView.Caret.Position.TextInsertionIndex OrElse _TextView.Caret.Position.TextInsertionIndex > formattedTextLines1.LastFullyVisibleLine.LineSpan.End Then
                _TextView.Selection.Clear()
                formattedTextLines1.FirstFullyVisibleLine.MoveCaretToLocation(_TextView.Caret.PreferredBounds.Left)
                _TextView.Caret.EnsureVisible()
            End If
        End Sub

        Public Sub TransposeCharacter(undoHistory1 As UndoHistory) Implements IEditorOperations.TransposeCharacter
            If undoHistory1 Is Nothing Then
                Throw New ArgumentNullException("undoHistory")
            End If

            Dim position1 As ICaretPosition = _textView.Caret.Position
            Dim lineFromPosition As ITextSnapshotLine = _textView.TextSnapshot.GetLineFromPosition(position1.TextInsertionIndex)
            Dim text As String = lineFromPosition.GetText()

            If StringInfo.ParseCombiningCharacters(text).Length >= 2 Then
                Dim textElementSpan As Span
                Dim textElementSpan2 As Span
                If position1.TextInsertionIndex = lineFromPosition.Start Then
                    textElementSpan = _TextView.GetTextElementSpan(position1.TextInsertionIndex)
                    textElementSpan2 = _TextView.GetTextElementSpan(textElementSpan.End)
                ElseIf position1.TextInsertionIndex = lineFromPosition.End Then
                    textElementSpan2 = _TextView.GetTextElementSpan(position1.TextInsertionIndex - 1)
                    textElementSpan = _TextView.GetTextElementSpan(textElementSpan2.Start - 1)
                Else
                    textElementSpan2 = _TextView.GetTextElementSpan(position1.TextInsertionIndex)
                    textElementSpan = _TextView.GetTextElementSpan(textElementSpan2.Start - 1)
                End If

                Dim text2 As String = _TextView.TextSnapshot.GetText(textElementSpan.Start, textElementSpan.Length)
                Dim text3 As String = _TextView.TextSnapshot.GetText(textElementSpan2.Start, textElementSpan2.Length)
                Dim start1 As Integer = textElementSpan.Start
                Dim text4 As String = text3 & text2

                Using undoTransaction1 As UndoTransaction = undoHistory1.CreateTransaction("Transpose Character")
                    ReplaceText(New Span(start1, text4.Length), text4, undoHistory1)
                    undoTransaction1.Complete()
                End Using
                _TextView.Caret.EnsureVisible()
                _TextView.Caret.CapturePreferredBounds()
            End If

        End Sub

        Public Sub TransposeLine(undoHistory1 As UndoHistory) Implements IEditorOperations.TransposeLine
            If undoHistory1 Is Nothing Then
                Throw New ArgumentNullException("undoHistory")
            End If

            If _textView.TextSnapshot.LineCount < 2 Then
                Return
            End If

            Dim position1 As ICaretPosition = _textView.Caret.Position
            Dim lineFromPosition As ITextSnapshotLine = _textView.TextSnapshot.GetLineFromPosition(position1.TextInsertionIndex)
            _textView.Caret.EnsureVisible()
            Dim formattedTextLines1 As IFormattedTextLineCollection = _textView.FormattedTextLines
            Dim textElementIndex As Integer = formattedTextLines1.GetTextLineContainingPosition(position1.TextInsertionIndex).GetTextElementIndex(position1.TextInsertionIndex)
            Dim lineNumber1 As Integer = lineFromPosition.LineNumber
            Dim textSnapshotLine As ITextSnapshotLine
            Dim textSnapshotLine2 As ITextSnapshotLine
            Dim characterIndex1 As Integer

            If lineNumber1 = _textView.TextSnapshot.LineCount - 1 Then
                textSnapshotLine = _textView.TextSnapshot.GetLineFromLineNumber(lineNumber1 - 1)
                textSnapshotLine2 = lineFromPosition
                Dim textLineContainingPosition As ITextLine = formattedTextLines1.GetTextLineContainingPosition(textSnapshotLine.Start)
                Dim textElementIndex2 As Integer = textLineContainingPosition.GetTextElementIndex(textSnapshotLine.End)
                characterIndex1 = (If((textElementIndex < textElementIndex2), (textLineContainingPosition.GetTextElementSpan(textElementIndex).Start + textSnapshotLine.LineBreakLength + textSnapshotLine2.Length), _textView.TextSnapshot.Length))
            Else
                textSnapshotLine = lineFromPosition
                textSnapshotLine2 = _textView.TextSnapshot.GetLineFromLineNumber(lineNumber1 + 1)
                characterIndex1 = position1.TextInsertionIndex + textSnapshotLine.LineBreakLength + textSnapshotLine2.Length
            End If

            Using undoTransaction1 As UndoTransaction = undoHistory1.CreateTransaction("Transpose Line")
                Using textEdit As ITextEdit = _textView.TextBuffer.CreateEdit()
                    If textEdit.Replace(textSnapshotLine.Extent, textSnapshotLine2.GetText()) AndAlso textEdit.Replace(textSnapshotLine2.Extent, textSnapshotLine.GetText()) Then
                        Dim undo As New BeforeTextBufferChangeUndoPrimitive(_textView, position1.CharacterIndex, position1.Placement, _textView.Selection.ActiveSnapshotSpan)
                        undoHistory1.CurrentTransaction.AddUndo(undo)
                        textEdit.Apply()
                        _textView.Caret.MoveTo(characterIndex1)
                        _textView.Selection.Clear()
                        Dim undo2 As New AfterTextBufferChangeUndoPrimitive(_textView, _textView.Caret.Position.CharacterIndex, _textView.Caret.Position.Placement, _textView.Selection.ActiveSnapshotSpan)
                        undoHistory1.CurrentTransaction.AddUndo(undo2)
                    End If
                    undoTransaction1.Complete()
                End Using
            End Using

            _textView.Caret.EnsureVisible()
            _textView.Caret.CapturePreferredBounds()
        End Sub

        Private Sub ChangeCase(letterCase As LetterCase, undoHistory As UndoHistory)
            If undoHistory Is Nothing Then
                Throw New ArgumentNullException("undoHistory")
            End If

            If _TextView.Caret.Position.TextInsertionIndex = _TextView.TextSnapshot.Length AndAlso _TextView.Selection.IsEmpty Then
                Return
            End If

            Dim position1 As ICaretPosition = _TextView.Caret.Position
            If Not _TextView.Selection.IsEmpty Then
                Dim start1 As Integer = _TextView.Selection.ActiveSnapshotSpan.Start
                Dim length1 As Integer = _TextView.Selection.ActiveSnapshotSpan.Length
                Dim text As String = _TextView.TextSnapshot.GetText(start1, length1)
                Dim text2 As String = (If((letterCase <> 0), text.ToLower(CultureInfo.CurrentCulture), text.ToUpper(CultureInfo.CurrentCulture)))

                Using undoTransaction = undoHistory.CreateTransaction("Make " & (If((letterCase = LetterCase.Uppercase), "Uppercase", "Lowercase")))
                    ReplaceText(_TextView.Selection.ActiveSnapshotSpan.Span, text2, TextEditAction.Replace, undoHistory, preserveCaretAndSelection:=True)
                    undoTransaction.Complete()
                End Using

            Else
                Dim lineFromPosition As ITextSnapshotLine = _TextView.TextSnapshot.GetLineFromPosition(position1.TextInsertionIndex)
                If position1.TextInsertionIndex <> lineFromPosition.End Then
                    Dim text As String = _TextView.TextSnapshot.GetText(position1.TextInsertionIndex, _TextView.GetTextElementSpan(position1.TextInsertionIndex).Length)
                    Dim text2 As String = (If((letterCase <> 0), text.ToLower(CultureInfo.CurrentCulture), text.ToUpper(CultureInfo.CurrentCulture)))

                    Using undoTransaction = undoHistory.CreateTransaction("Make " & (If((letterCase = LetterCase.Uppercase), "Uppercase", "Lowercase")))
                        ReplaceText(New Span(position1.TextInsertionIndex, text2.Length), text2, undoHistory)
                        undoTransaction.Complete()
                    End Using

                Else
                    MoveCaret(CaretMovementDirection.Next, extendSelection:=False)
                End If
            End If

            _TextView.Caret.EnsureVisible()
        End Sub

        Public Sub MakeUppercase(undoHistory1 As UndoHistory) Implements IEditorOperations.MakeUppercase
            ChangeCase(LetterCase.Uppercase, undoHistory1)
        End Sub

        Public Sub MakeLowercase(undoHistory As UndoHistory) Implements IEditorOperations.MakeLowercase
            ChangeCase(LetterCase.Lowercase, undoHistory)
        End Sub

        Private Function ColumnOffsetOfPositionInLine(position1 As Integer, line As ITextSnapshotLine) As Integer
            Return position1 - line.Start
        End Function

        Private Function ReplaceSelection(text As String, textEditAction As TextEditAction, undoHistory As UndoHistory) As Boolean
            If undoHistory Is Nothing Then
                Throw New ArgumentNullException("undoHistory")
            End If

            Dim s = If(_TextView.Selection.IsEmpty, New Span(_TextView.Caret.Position.TextInsertionIndex, 0), _TextView.Selection.ActiveSnapshotSpan)
            Return ReplaceText(s, text, textEditAction, undoHistory)
        End Function

        Private Function ReplaceText(span As Span, text As String, textEditAction As TextEditAction, undoHistory As UndoHistory) As Boolean
            Return ReplaceText(span, text, textEditAction, undoHistory, preserveCaretAndSelection:=False)
        End Function

        Private Function ReplaceText(span As Span, text As String, textEditAction As TextEditAction, undoHistory As UndoHistory, preserveCaretAndSelection As Boolean) As Boolean
            EditorTrace.TraceTextReplaceStart()

            Using textEdit As ITextEdit = _TextView.TextBuffer.CreateEdit()
                If If(span.Length = 0, textEdit.CanInsert(span.Start), textEdit.CanDeleteOrReplace(span)) Then
                    Dim caretPos = _TextView.Caret.Position
                    Dim selectedSpan = _TextView.Selection.ActiveSnapshotSpan
                    undoHistory.CurrentTransaction.AddUndo(
                            New BeforeTextBufferChangeUndoPrimitive(
                               _TextView,
                               caretPos.CharacterIndex,
                               caretPos.Placement,
                               selectedSpan)
                    )
                    textEdit.Replace(span, text)
                    textEdit.Apply()

                    If preserveCaretAndSelection Then
                        _TextView.Caret.MoveTo(caretPos.CharacterIndex, caretPos.Placement)
                        _TextView.Selection.ActiveSnapshotSpan = New SnapshotSpan(_TextView.TextSnapshot, selectedSpan)
                    Else
                        ResetSelection()
                    End If

                    _TextView.Caret.CapturePreferredBounds()
                    undoHistory.CurrentTransaction.AddUndo(
                            New AfterTextBufferChangeUndoPrimitive(
                                _TextView,
                                _TextView.Caret.Position.CharacterIndex,
                                _TextView.Caret.Position.Placement,
                                _TextView.Selection.ActiveSnapshotSpan)
                    )
                    EditorTrace.TraceTextReplaceEnd()
                    Return True
                End If
            End Using

            Return False
        End Function

        Private Function InsertTextOverwriteMode(index As Integer, text As String, undoHistory As UndoHistory) As Boolean
            Dim text2 As String = ""
            Dim length As Integer = 0

            If _TextView.Caret.Position.TextInsertionIndex < _TextView.TextSnapshot.GetLineFromPosition(_TextView.Caret.Position.TextInsertionIndex).End Then
                length = text.Length
                text2 = _TextView.TextSnapshot.GetText(_TextView.Caret.Position.TextInsertionIndex, length)
            End If

            Using textEdit = _TextView.TextBuffer.CreateEdit()
                If textEdit.CanDeleteOrReplace(New Span(index, text2.Length)) Then
                    undoHistory.CurrentTransaction.AddUndo(
                           New BeforeTextBufferChangeUndoPrimitive(
                               _TextView,
                               _TextView.Caret.Position.CharacterIndex,
                               _TextView.Caret.Position.Placement,
                               _TextView.Selection.ActiveSnapshotSpan
                          )
                    )
                    textEdit.Replace(index, text2.Length, text)
                    textEdit.Apply()
                    ResetSelection()
                    _TextView.Caret.CapturePreferredBounds()
                    undoHistory.CurrentTransaction.AddUndo(
                           New AfterTextBufferChangeUndoPrimitive(
                               _TextView,
                               _TextView.Caret.Position.CharacterIndex,
                               _TextView.Caret.Position.Placement,
                               _TextView.Selection.ActiveSnapshotSpan)
                    )
                    Return True
                End If

                Return False
            End Using
        End Function

        Private Function SkipWhitespaceForward(position1 As Integer) As Integer
            Dim textSnapshot1 As ITextSnapshot = _textView.TextSnapshot
            While position1 < _textView.TextSnapshot.Length
                Dim extentOfWord As TextExtent = _textStructureNavigator.GetExtentOfWord(New SnapshotPoint(textSnapshot1, position1))

                If extentOfWord.IsSignificant Then Exit While

                Dim [end] As Integer = _textView.TextSnapshot.GetLineFromPosition(extentOfWord.Span.Start).End
                If [end] <= extentOfWord.Span.End Then Return [end]

                position1 = extentOfWord.Span.End
            End While

            Return position1
        End Function

        Private Function SkipWhitespaceBackward(position1 As Integer) As Integer
            Dim textSnapshot1 As ITextSnapshot = _textView.TextSnapshot
            While position1 > 0
                Dim extentOfWord As TextExtent = _textStructureNavigator.GetExtentOfWord(New SnapshotPoint(textSnapshot1, position1))
                If extentOfWord.IsSignificant Then
                    Return extentOfWord.Span.Start
                End If

                Dim start1 As Integer = textSnapshot1.GetLineFromPosition(position1).Start
                If start1 >= extentOfWord.Span.Start Then Return start1

                position1 = extentOfWord.Span.Start - 1
            End While

            Return position1
        End Function

        Private Function GetNextStartOfWord(currentPosition As Integer) As Integer
            Dim textSnapshot1 As ITextSnapshot = _textView.TextSnapshot
            If currentPosition >= textSnapshot1.Length - 1 Then
                Return textSnapshot1.Length
            End If

            Return SkipWhitespaceForward(_textStructureNavigator.GetExtentOfWord(New SnapshotPoint(textSnapshot1, currentPosition)).Span.End)
        End Function

        Private Function GetPreviousStartOfWord(currentPosition As Integer) As Integer
            If currentPosition = 0 Then Return 0

            Dim extentOfWord As TextExtent = _textStructureNavigator.GetExtentOfWord(New SnapshotPoint(_textView.TextSnapshot, currentPosition))
            Dim num As Integer = 0

            If extentOfWord.IsSignificant AndAlso extentOfWord.Span.Start < currentPosition Then
                Return extentOfWord.Span.Start
            End If

            If extentOfWord.Span.Start > 0 Then
                Return SkipWhitespaceBackward(extentOfWord.Span.Start - 1)
            End If

            Return extentOfWord.Span.Start
        End Function

        Private Function GetPreviousCodePoint(currentIndex As Integer) As Integer
            Dim textElementSpan As Span = _textView.GetTextElementSpan(currentIndex)
            Dim num As Integer = 0

            num = (If((textElementSpan.Start <= 0), textElementSpan.Start, _textView.GetTextElementSpan(textElementSpan.Start - 1).Start))
            If currentIndex > 1 Then
                Dim text As String = _textView.TextSnapshot.GetText(currentIndex - 2, 2)
                If text = Microsoft.VisualBasic.vbNewLine Then
                    num -= 2
                End If
            End If

            Return num
        End Function

        Private Function GetNextCodePoint(currentIndex As Integer) As Integer
            Return _textView.GetTextElementSpan(currentIndex).End
        End Function

        Private Sub MoveCharacter(direction As CaretMovementDirection, [select] As Boolean)
            MoveCaret(direction, [select])
            If Not [select] Then _TextView.Selection.Clear()

            _TextView.Caret.EnsureVisible()
            _textView.Caret.CapturePreferredBounds()
        End Sub

        Private Function MoveCaret(direction As CaretMovementDirection, extendSelection As Boolean) As ICaretPosition
            Dim position1 As ICaretPosition = _textView.Caret.Position
            _textView.ViewScroller.EnsureSpanVisible(New Span(position1.CharacterIndex, 0), 0.0, 0.0)
            Dim span1 As Span = _textView.Selection.ActiveSnapshotSpan
            Dim num As Integer = (If(_textView.Selection.IsEmpty, position1.TextInsertionIndex, (If((Not _textView.Selection.IsActiveSpanReversed), span1.Start, span1.End))))
            Dim caretPosition As ICaretPosition = (If((direction = CaretMovementDirection.Previous), (If((Not _textView.Selection.IsEmpty AndAlso Not extendSelection), _textView.Caret.MoveTo(span1.Start), _textView.Caret.MoveToPreviousCaretPosition())), (If((Not _textView.Selection.IsEmpty AndAlso Not extendSelection), _textView.Caret.MoveTo(span1.End), _textView.Caret.MoveToNextCaretPosition()))))

            If extendSelection Then
                If num < caretPosition.TextInsertionIndex Then
                    _TextView.Selection.ActiveSnapshotSpan = New SnapshotSpan(_TextView.TextSnapshot, Span.FromBounds(num, caretPosition.TextInsertionIndex))
                    _TextView.Selection.IsActiveSpanReversed = False
                Else
                    _TextView.Selection.ActiveSnapshotSpan = New SnapshotSpan(_TextView.TextSnapshot, Span.FromBounds(caretPosition.TextInsertionIndex, num))
                    _TextView.Selection.IsActiveSpanReversed = True
                End If
            End If

            Return caretPosition
        End Function

        Private Function GetExtentOfCurrentWord(currentPosition As SnapshotPoint) As TextExtent
            Dim extentOfWord = _textStructureNavigator.GetExtentOfWord(currentPosition)
            If extentOfWord.IsSignificant Then Return extentOfWord

            Dim containingLine = currentPosition.GetContainingLine()
            Dim start = Math.Max(containingLine.Start, extentOfWord.Span.Start)
            Dim [end] = Math.Min(containingLine.End, extentOfWord.Span.End)
            Dim length = [end] - start

            Select Case length
                Case 0
                    If start = containingLine.Start Then
                        Return New TextExtent(currentPosition.Position, 0, isSignificant:=False)
                    End If
                    Return GetExtentOfCurrentWord(New SnapshotPoint(currentPosition.Snapshot, start - 1))

                Case 1
                    If start = containingLine.Start Then
                        Return New TextExtent(start, 1, isSignificant:=False)
                    End If
                    Return GetExtentOfCurrentWord(New SnapshotPoint(currentPosition.Snapshot, start - 1))

                Case Else
                    Return New TextExtent(start, length, isSignificant:=False)
            End Select

        End Function

        Public Sub [Select](start As Integer, length As Integer) Implements IEditorOperations.Select
            _TextView.Caret.MoveTo(start + length)
            _TextView.Caret.EnsureVisible()
            _TextView.Selection.ActiveSnapshotSpan = New SnapshotSpan(_TextView.TextSnapshot, New Span(start, length))
        End Sub
    End Class
End Namespace
