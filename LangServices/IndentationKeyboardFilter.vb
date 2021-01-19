Imports System.ComponentModel.Composition
Imports System.Windows.Input
Imports Microsoft.Nautilus.Core.Undo
Imports Microsoft.Nautilus.Text
Imports Microsoft.Nautilus.Text.Editor
Imports Microsoft.Nautilus.Text.Operations
Imports Microsoft.SmallBasic.Completion

Namespace Microsoft.SmallBasic.LanguageService
    Public Class IndentationKeyboardFilter
        Inherits KeyboardFilter

        Public completionHelper As CompletionHelper = New CompletionHelper()
        <Import>
        Public Property EditorOperationsProvider As IEditorOperationsProvider
        <Import>
        Public Property UndoHistoryRegistry As IUndoHistoryRegistry

        Public Overrides Sub KeyDown(ByVal textView As IAvalonTextView, ByVal args As KeyEventArgs)
            If args.Key = Key.Return Then
                Dim textInsertionIndex = textView.Caret.Position.TextInsertionIndex
                Dim lineNumberFromPosition = textView.TextSnapshot.GetLineNumberFromPosition(textInsertionIndex)
                UpdateIndentation(textView.TextSnapshot, lineNumberFromPosition)
                AddHandler textView.TextBuffer.Changed, AddressOf OnTextBufferChanged
            End If

            MyBase.KeyDown(textView, args)
        End Sub

        Private Sub OnTextBufferChanged(ByVal sender As Object, ByVal e As TextChangedEventArgs)
            Dim textBuffer As ITextBuffer = TryCast(sender, ITextBuffer)
            RemoveHandler textBuffer.Changed, AddressOf OnTextBufferChanged

            If e.Changes(CInt(0)).NewText.StartsWith(VisualBasic.Constants.vbCrLf) Then
                Dim after = e.After
                Dim lineNumberFromPosition = after.GetLineNumberFromPosition(e.Changes(0).NewEnd)
                UpdateIndentation(after, lineNumberFromPosition)
            End If
        End Sub

        Public Sub UpdateIndentation(ByVal snapshot As ITextSnapshot, ByVal lineNumber As Integer)
            If lineNumber <> 0 Then
                Dim source As TextBufferReader = New TextBufferReader(snapshot)
                Dim indentationLevel = completionHelper.GetIndentationLevel(source, lineNumber)
                Dim indentationLevel2 = completionHelper.GetIndentationLevel(lineNumber - 1)
                Dim lineFromLineNumber = snapshot.GetLineFromLineNumber(lineNumber)
                Dim lineFromLineNumber2 = snapshot.GetLineFromLineNumber(lineNumber - 1)
                Dim positionOfNextNonWhiteSpaceCharacter = lineFromLineNumber.GetPositionOfNextNonWhiteSpaceCharacter(0)
                Dim positionOfNextNonWhiteSpaceCharacter2 = lineFromLineNumber2.GetPositionOfNextNonWhiteSpaceCharacter(0)
                Dim num = positionOfNextNonWhiteSpaceCharacter2

                If indentationLevel > indentationLevel2 Then
                    num += 2
                ElseIf indentationLevel < indentationLevel2 Then
                    num -= 2
                End If

                If num < 0 Then
                    num = 0
                End If

                If positionOfNextNonWhiteSpaceCharacter > num Then
                    snapshot.TextBuffer.Delete(New Span(lineFromLineNumber.Start, positionOfNextNonWhiteSpaceCharacter - num))
                ElseIf positionOfNextNonWhiteSpaceCharacter < num Then
                    snapshot.TextBuffer.Insert(lineFromLineNumber.Start, New String(" "c, num - positionOfNextNonWhiteSpaceCharacter))
                End If
            End If
        End Sub
    End Class
End Namespace
