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
            Dim iden = Indentation.CalculateIndentation(snapshot, lineNumber)
            Dim line = snapshot.GetLineFromLineNumber(lineNumber)
            Dim positionOfNextToken = line.GetPositionOfNextNonWhiteSpaceCharacter(0)

            If positionOfNextToken > iden Then
                snapshot.TextBuffer.Delete(New Span(line.Start, positionOfNextToken - iden))
            ElseIf positionOfNextToken < iden Then
                snapshot.TextBuffer.Insert(line.Start, New String(" "c, iden - positionOfNextToken))
            End If

        End Sub

    End Class

    Public NotInheritable Class Indentation
        Public Shared Function CalculateIndentation(snapshot As ITextSnapshot, lineNumber As Integer) As Integer
            Dim iden As Integer
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
                    If prevLineText <> "" Then Exit Do
                    n -= 1
                Loop

                If prevLineTrimmedText <> "" Then
                    Dim prevIndentationLevel = prevLineText.Length - prevLineText.TrimStart(" "c, vbTab).Length
                    Dim prevTokens = New LineScanner().GetTokenList(prevLineText, 1)
                    Dim tokens = New LineScanner().GetTokenList(lineText, 1)

                    Select Case prevTokens.Current.Token
                        Case Token.For, Token.If, Token.Sub, Token.While
                            iden = prevIndentationLevel + 3
                        Case Else
                            Select Case tokens.Current.Token
                                Case Token.EndFor, Token.EndIf, Token.EndSub, Token.EndWhile
                                    iden = prevIndentationLevel - 3
                                Case Else
                                    iden = prevIndentationLevel
                            End Select
                    End Select

                    If iden < 0 Then iden = 0
                Else
                    iden = indentationLevel
                End If
            Else
                iden = indentationLevel
            End If
            Return iden
        End Function

    End Class
End Namespace
