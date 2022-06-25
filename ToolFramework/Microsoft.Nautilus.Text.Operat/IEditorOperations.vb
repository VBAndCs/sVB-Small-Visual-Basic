Imports Microsoft.Nautilus.Core.Undo
Imports Microsoft.Nautilus.Text.Editor

Namespace Microsoft.Nautilus.Text.Operations
    Public Interface IEditorOperations
        ReadOnly Property CanPaste As Boolean

        ReadOnly Property TextView As ITextView

        Property ConvertTabsToSpace As Boolean

        Property TabSize As Integer

        Property OverwriteMode As Boolean

        Sub MoveCharacterRight([select] As Boolean)

        Sub MoveCharacterLeft([select] As Boolean)

        Sub MoveToNextWord([select] As Boolean)

        Sub MoveToPreviousWord([select] As Boolean)

        Sub MoveLineUp([select] As Boolean)

        Sub MoveLineDown([select] As Boolean)

        Sub PageUp([select] As Boolean)

        Sub PageDown([select] As Boolean)

        Sub MoveToEndOfLine([select] As Boolean)

        Sub MoveToStartOfLine([select] As Boolean)

        Sub GotoLine(lineNumber As Integer)

        Sub MoveToStartOfDocument([select] As Boolean)

        Sub MoveToEndOfDocument([select] As Boolean)

        Sub MoveCurrentLineToTop()

        Sub MoveCurrentLineToBottom()

        Sub ScrollUpAndMoveCaretIfNecessary()

        Sub ScrollDownAndMoveCaretIfNecessary()

        Sub DeleteCharacterToLeft(undoHistory1 As UndoHistory)

        Sub DeleteCharacterToRight(undoHistory1 As UndoHistory)

        Sub DeleteWordToRight(undoHistory1 As UndoHistory)

        Sub DeleteWordToLeft(undoHistory1 As UndoHistory)

        Sub YankCurrentLine(undoHistory1 As UndoHistory)

        Sub InsertNewline(undoHistory1 As UndoHistory)

        Sub InsertTab(undoHistory1 As UndoHistory)

        Sub RemovePreviousTab(undoHistory1 As UndoHistory)

        Sub IndentSelection(undoHistory1 As UndoHistory)

        Sub UnindentSelection(undoHistory1 As UndoHistory)

        Sub InsertText(text1 As String, undoHistory1 As UndoHistory)

        Function DeleteSelection(undoHistory1 As UndoHistory) As Boolean

        Sub ReplaceSelection(text1 As String, undoHistory1 As UndoHistory)

        Sub TransposeCharacter(undoHistory1 As UndoHistory)

        Sub TransposeLine(undoHistory1 As UndoHistory)

        Sub MakeLowercase(undoHistory1 As UndoHistory)

        Sub MakeUppercase(undoHistory1 As UndoHistory)

        Sub ReplaceText(replaceSpan As Span, text1 As String, undoHistory1 As UndoHistory)

        Function ReplaceAllMatches(searchText As String, replaceText1 As String, matchCase As Boolean, matchWholeWord As Boolean, useRegularExpressions As Boolean, undoHistory1 As UndoHistory) As Integer

        Sub SelectCurrentWord()

        Sub SelectLine(lineNumber As Integer, extendSelection As Boolean)

        Sub SelectAll()

        Sub ExtendSelection(newEnd As Integer)

        Sub MoveCaretAndExtendSelection(textLine As ITextLine, horizontalOffset As Double)

        Sub ResetSelection()

        Sub CopySelection()

        Sub CutSelection(undoHistory1 As UndoHistory)

        Sub Paste(undoHistory1 As UndoHistory)
    End Interface
End Namespace
