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

        Public Overrides Sub KeyDown(textView As IAvalonTextView, args As KeyEventArgs)
            If args.Key = Key.Return Then
                Dim textInsertionIndex = textView.Caret.Position.TextInsertionIndex
                Dim lineNumber = textView.TextSnapshot.GetLineNumberFromPosition(textInsertionIndex)
                UpdateIndentation(textView.TextSnapshot, lineNumber)
                AddHandler textView.TextBuffer.Changed, AddressOf OnTextBufferChanged
            End If

            MyBase.KeyDown(textView, args)
        End Sub

        Private Sub OnTextBufferChanged(sender As Object, e As TextChangedEventArgs)
            Dim textBuffer = TryCast(sender, ITextBuffer)
            RemoveHandler textBuffer.Changed, AddressOf OnTextBufferChanged

            If e.Changes(0).NewText.StartsWith(vbCrLf) Then
                Dim after = e.After
                Dim lineNumber = after.GetLineNumberFromPosition(e.Changes(0).NewEnd)
                UpdateIndentation(after, lineNumber)
            End If
        End Sub

        Public Sub UpdateIndentation(snapshot As ITextSnapshot, lineNumber As Integer)
            Dim inden = Indentation.CalculateIndentation(snapshot, lineNumber)
            Dim line = snapshot.GetLineFromLineNumber(lineNumber)
            Dim positionOfNextToken = line.GetPositionOfNextNonWhiteSpaceCharacter(0)

            If positionOfNextToken > inden Then
                snapshot.TextBuffer.Delete(New Span(line.Start, positionOfNextToken - inden))
            ElseIf positionOfNextToken < inden Then
                snapshot.TextBuffer.Insert(line.Start, New String(" "c, inden - positionOfNextToken))
            End If

        End Sub

    End Class

    Public NotInheritable Class Indentation
        Public Shared Function CalculateIndentation(snapshot As ITextSnapshot, lineNumber As Integer) As Integer
            Dim inden As Integer
            Dim line = snapshot.GetLineFromLineNumber(lineNumber)
            Dim lineText = line.GetText()
            Dim indentationLevel = lineText.Length - lineText.TrimStart(" "c, vbTab).Length

            If lineNumber > 0 Then
                Dim positionOfNextToken = line.GetPositionOfNextNonWhiteSpaceCharacter(0)
                Dim prevLine As ITextSnapshotLine
                Dim n = lineNumber - 1
                Dim prevLineText = ""
                Dim prevLineTrimmedText = ""

                ' Find last non-empty line
                Do While n > -1
                    prevLine = snapshot.GetLineFromLineNumber(n)
                    prevLineText = prevLine.GetText()
                    prevLineTrimmedText = prevLineText.Trim(" "c, vbTab)
                    If prevLineTrimmedText <> "" Then Exit Do
                    n -= 1
                Loop

                If prevLineTrimmedText <> "" Then
                    Dim prevIndentationLevel = prevLineText.Length - prevLineText.TrimStart(" "c, vbTab).Length
                    Dim prevTokens = New LineScanner().GetTokenList(prevLineText, 1)
                    Dim tokens = New LineScanner().GetTokenList(lineText, 1)

                    Select Case tokens.Current.Token
                        Case Token.EndFor, Token.Next, Token.ElseIf, Token.Else, Token.EndIf, Token.EndSub, Token.EndWhile, Token.Wend
                            Select Case prevTokens.Current.Token
                                Case Token.For, Token.If, Token.ElseIf, Token.Else, Token.Sub, Token.While
                                    inden = prevIndentationLevel
                                Case Else
                                    inden = prevIndentationLevel - 3
                            End Select
                        Case Else
                                    Select Case prevTokens.Current.Token
                                        Case Token.For, Token.If, Token.ElseIf, Token.Else, Token.Sub, Token.While
                                            inden = prevIndentationLevel + 3
                                        Case Else
                                            inden = prevIndentationLevel
                                    End Select
                            End Select

                            If inden < 0 Then inden = 0
                Else
                    inden = indentationLevel
                End If
            Else
                inden = indentationLevel
            End If
            Return inden
        End Function

    End Class
End Namespace
