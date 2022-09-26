Imports System.ComponentModel.Composition
Imports System.Windows.Input
Imports Microsoft.Nautilus.Core.Undo
Imports Microsoft.Nautilus.Text
Imports Microsoft.Nautilus.Text.Editor
Imports Microsoft.Nautilus.Text.Operations

Namespace Microsoft.SmallVisualBasic.LanguageService
    Public Class CompletionKeyboardFilter
        Inherits KeyboardFilter

        <Import>
        Public Property EditorOperationsProvider As IEditorOperationsProvider
        <Import>
        Public Property UndoHistoryRegistry As IUndoHistoryRegistry



        Public Overrides Sub KeyDown(textView As IAvalonTextView, args As KeyEventArgs)
            Dim completionSurface = textView.Properties.GetProperty(Of CompletionAdornmentSurface)()

            If completionSurface IsNot Nothing AndAlso completionSurface.IsAdornmentVisible Then
                Dim completionListBox = completionSurface.CompletionListBox
                Dim __ = completionListBox.SelectedIndex

                Select Case args.Key
                    Case Key.LeftCtrl, Key.RightCtrl
                        completionSurface.FadeCompletionList()

                    Case Key.Escape
                        completionSurface.Adornment.Dismiss(force:=True)
                        textView.VisualElement.Focus()
                        args.Handled = True

                    Case Key.Up
                        completionListBox.MoveUp()
                        args.Handled = True

                    Case Key.Down
                        completionListBox.MoveDown()
                        args.Handled = True

                    Case Key.Return
                        args.Handled = CommitConditionally(textView, completionSurface)

                    Case Key.Space, Key.Tab
                        args.Handled = CommitConditionally(textView, completionSurface, " ")

                    Case Key.OemPeriod
                        CommitConditionally(textView, completionSurface)

                End Select

            Else
                Dim provider = textView.Properties.GetProperty(Of CompletionProvider)()
                If provider Is Nothing Then Return

                If args.Key = Key.Space AndAlso args.KeyboardDevice.Modifiers = ModifierKeys.Control Then
                    provider.ShowCompletionAdornment(textView.TextSnapshot, textView.Caret.Position.TextInsertionIndex, True)
                    args.Handled = True

                ElseIf args.Key = Key.F1 Then
                    provider.ShowHelp(True)
                End If
            End If
        End Sub

        Public Overrides Sub KeyUp(textView As IAvalonTextView, args As KeyEventArgs)
            Dim adornmentSurface As CompletionAdornmentSurface = Nothing

            If (args.Key = Key.LeftCtrl OrElse args.Key = Key.RightCtrl) AndAlso
                        textView.Properties.TryGetProperty(Of CompletionAdornmentSurface)(adornmentSurface) AndAlso
                        adornmentSurface.IsAdornmentVisible Then
                adornmentSurface.UnfadeCompletionList()
            End If

            MyBase.KeyUp(textView, args)
        End Sub

        Public Overrides Sub TextInput(textView As IAvalonTextView, args As TextCompositionEventArgs)
            Dim completionSurface As CompletionAdornmentSurface = Nothing

            If textView.Properties.TryGetProperty(GetType(CompletionAdornmentSurface), completionSurface) Then
                If completionSurface.IsAdornmentVisible Then
                    Select Case args.Text
                        Case "+", "-", "*", "/"
                            args.Handled = CommitConditionally(textView, completionSurface, " " & args.Text & " ")

                        Case "="
                            Dim pos = textView.Caret.Position.CharacterIndex
                            Select Case textView.TextSnapshot.GetText(pos - 1, 1)
                                Case ">", "<"

                                Case Else
                                    args.Handled = CommitConditionally(textView, completionSurface, " " & args.Text & " ", True)
                            End Select

                        Case "!", ")", "[", "]", "{", "}"
                            CommitConditionally(textView, completionSurface)

                        Case ",", "("
                            args.Handled = CommitConditionally(textView, completionSurface, args.Text & " ")

                    End Select
                End If
            End If
            MyBase.TextInputUpdate(textView, args)
        End Sub

        Private Function CommitConditionally(
                                textView As IAvalonTextView,
                                completionSurface As CompletionAdornmentSurface,
                                Optional extraText As String = "",
                                Optional showCompletionAdornmentAgain As Boolean = False
                       ) As Boolean

            If Not completionSurface.IsFaded AndAlso completionSurface.CompletionListBox.SelectedItem IsNot Nothing Then
                Dim editorOperations = EditorOperationsProvider.GetEditorOperations(textView)
                Dim compList = completionSurface.CompletionListBox
                Dim itemWrapper = CType(compList.SelectedItem, CompletionItemWrapper)
                Dim item = itemWrapper.CompletionItem
                Dim repWith = item.ReplacementText
                Dim provider = textView.Properties.GetProperty(Of CompletionProvider)()
                Dim replaceSpan = provider.GetReplacementSpane()

                Dim key = item.HistoryKey
                If key <> "" Then CompletionProvider.compHistory(key) = item.DisplayName

                If extraText.EndsWith("( ") AndAlso (repWith.EndsWith("(") OrElse repWith.EndsWith(")")) Then
                    extraText = ""
                End If

                editorOperations.ReplaceText(
                    replaceSpan,
                    repWith & extraText,
                    UndoHistoryRegistry.GetHistory(textView.TextBuffer))

                If completionSurface.Adornment IsNot Nothing Then
                    completionSurface.Adornment.Dismiss(force:=False)
                End If

                textView.VisualElement.Focus()

                If extraText = ", " OrElse repWith.EndsWith("(") Then
                    provider.ShowHelp(True)
                ElseIf showCompletionAdornmentAgain Then
                    provider.ShowCompletionAdornment(textView.TextSnapshot, textView.Caret.Position.TextInsertionIndex, True)
                End If
                Return True

            Else
                If completionSurface.Adornment IsNot Nothing Then
                    completionSurface.Adornment.Dismiss(force:=False)
                End If
                textView.VisualElement.Focus()
                Return False
            End If

        End Function

        Private Function CommitItem(textView As IAvalonTextView, adornmentSurface As CompletionAdornmentSurface) As Boolean
            Dim compList = adornmentSurface.CompletionListBox
            Dim itemWrapper = TryCast(compList.SelectedItem, CompletionItemWrapper)
            Dim provider As CompletionProvider = Nothing

            If itemWrapper IsNot Nothing Then
                If textView.Properties.TryGetProperty(GetType(CompletionProvider), provider) Then
                    provider.CommitItem(itemWrapper.CompletionItem)
                End If
                Return True

            ElseIf adornmentSurface.Adornment IsNot Nothing Then
                adornmentSurface.Adornment.Dismiss(force:=False)
            End If

            Return False
        End Function
    End Class
End Namespace
