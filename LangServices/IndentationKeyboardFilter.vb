Imports System.ComponentModel.Composition
Imports System.Windows.Input
Imports System.Windows.Threading
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
                Dim after = e.After
                Dim lineNumber = after.GetLineNumberFromPosition(e.Changes(0).NewEnd)
                CompilerService.FormatDocument(textBuffer, lineNumber)
            End If
        End Sub

    End Class

End Namespace
