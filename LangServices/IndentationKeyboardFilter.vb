Imports System.ComponentModel.Composition
Imports System.Windows.Threading
Imports Microsoft.Nautilus.Core.Undo
Imports Microsoft.Nautilus.Text
Imports Microsoft.Nautilus.Text.Editor
Imports Microsoft.Nautilus.Text.Operations
Imports Microsoft.SmallVisualBasic.Completion

Namespace Microsoft.SmallVisualBasic.LanguageService
    Public Class IndentationKeyboardFilter
        Inherits KeyboardFilter

        Public completionHelper As CompletionHelper = New CompletionHelper()

        <Import>
        Public Property EditorOperationsProvider As IEditorOperationsProvider

        <Import>
        Public Property UndoHistoryRegistry As IUndoHistoryRegistry

        Dim textView As ITextView
        Public Overrides Sub KeyDown(textView As IAvalonTextView, args As KeyEventArgs)
            Me.textView = textView
            If args.Key = Key.Return Then
                textView.VisualElement.Dispatcher.Invoke(DispatcherPriority.Render,
                     Sub()
                         AddHandler textView.TextBuffer.Changed, AddressOf OnTextBufferChanged
                     End Sub)
            End If

            MyBase.KeyDown(textView, args)
        End Sub

        Private Sub OnTextBufferChanged(sender As Object, e As TextChangedEventArgs)
            Dim textBuffer = TryCast(sender, ITextBuffer)
            RemoveHandler textBuffer.Changed, AddressOf OnTextBufferChanged

            If e.Changes(0).NewText.StartsWith(vbCrLf) Then
                Dim snapshot = e.After
                Dim lineNumber = snapshot.GetLineNumberFromPosition(e.Changes(0).NewEnd)
                SpliteStringLiteral(textBuffer, lineNumber)
                FormatDocument(textBuffer, lineNumber)
            End If
        End Sub

        Const _QUOTE = ChrW(34)

        Sub SpliteStringLiteral(textBuffer As ITextBuffer, lineNumber As Integer)
            Dim snapshot = textBuffer.CurrentSnapshot
            Dim insertionIndex = textView.Caret.Position.TextInsertionIndex
            If insertionIndex = 0 Then Return

            Dim line1 = snapshot.Lines(lineNumber - 1)
            Dim line2 = snapshot.Lines(lineNumber)
            Dim tokens = LineScanner.GetTokens(line1.GetText(), line1.LineNumber)

            If tokens.Count > 0 Then
                Dim token = tokens.Last
                If token.Type = TokenType.StringLiteral Then
                    If token.Text.Length > 1 AndAlso token.Text.EndsWith(_QUOTE) Then Return

                    Using textEdit = textBuffer.CreateEdit()
                        Dim nl = _QUOTE & " & Text.NewLine &"
                        textEdit.Insert(line1.Start + token.EndColumn, nl)
                        tokens = LineScanner.GetTokens(line2.GetText(), line2.LineNumber)

                        If tokens.Count = 0 Then
                            textEdit.Insert(line2.End, _QUOTE & _QUOTE)
                            insertionIndex = line2.End + nl.Length + 1
                        Else
                            token = tokens(0)
                            textEdit.Insert(line2.Start + token.Column, _QUOTE)
                            insertionIndex = line2.Start + token.Column + nl.Length + 1
                        End If

                        textEdit.Apply()
                        textView.Caret.MoveTo(insertionIndex)
                    End Using
                End If
            End If
        End Sub

    End Class

End Namespace
