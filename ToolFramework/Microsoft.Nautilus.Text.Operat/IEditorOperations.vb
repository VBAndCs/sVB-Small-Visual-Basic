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

        Sub DeleteCharacterToLeft(undoHistory As UndoHistory)

        Sub DeleteCharacterToRight(undoHistory As UndoHistory)

        Sub DeleteWordToRight(undoHistory As UndoHistory)

        Sub DeleteWordToLeft(undoHistory As UndoHistory)

        Sub YankCurrentLine(undoHistory As UndoHistory)

        Sub InsertNewline(undoHistory As UndoHistory)

        Sub InsertTab(undoHistory As UndoHistory)

        Sub RemovePreviousTab(undoHistory As UndoHistory)

        Sub IndentSelection(undoHistory As UndoHistory)

        Sub UnindentSelection(undoHistory As UndoHistory)

        Sub InsertText(text As String, undoHistory As UndoHistory)

        Function DeleteSelection(undoHistory As UndoHistory) As Boolean

        Sub ReplaceSelection(text As String, undoHistory As UndoHistory)

        Sub TransposeCharacter(undoHistory As UndoHistory)

        Sub TransposeLine(undoHistory As UndoHistory)

        Sub MakeLowercase(undoHistory As UndoHistory)

        Sub MakeUppercase(undoHistory As UndoHistory)

        Sub ReplaceText(replaceSpan As Span, text As String, undoHistory As UndoHistory)

        Function ReplaceAllMatches(searchText As String, replacetext As String, matchCase As Boolean, matchWholeWord As Boolean, useRegularExpressions As Boolean, undoHistory As UndoHistory) As Integer

        Sub SelectCurrentWord()

        Sub SelectLine(lineNumber As Integer, extendSelection As Boolean)

        Sub SelectAll()

        Sub ExtendSelection(newEnd As Integer)

        Sub MoveCaretAndExtendSelection(textLine As ITextLine, horizontalOffset As Double)

        Sub ResetSelection()

        Sub CopySelection()

        Sub CutSelection(undoHistory As UndoHistory)

        Sub Paste(undoHistory As UndoHistory)

        Function GetCurrentWordSpan() As Span

        Function GetCurrentWord() As String
        Sub [Select](start As Integer, length As Integer)
    End Interface
End Namespace
