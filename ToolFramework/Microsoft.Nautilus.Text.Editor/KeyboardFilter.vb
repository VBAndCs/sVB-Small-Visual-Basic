Imports System.ComponentModel.Composition
Imports System.Windows.Input

Namespace Microsoft.Nautilus.Text.Editor
    <Export>
    Public MustInherit Class KeyboardFilter
        Public Overridable ReadOnly Property IsApplicableToHandledEvents As Boolean = False

        Public Overridable Sub Activate()
        End Sub

        Public Overridable Sub Deactivate()
        End Sub

        Public Overridable Sub GotKeyboardFocus(textView As IAvalonTextView, args As KeyboardFocusChangedEventArgs)
        End Sub

        Public Overridable Sub LostKeyboardFocus(textView As IAvalonTextView, args As KeyboardFocusChangedEventArgs)
        End Sub

        Public Overridable Sub KeyDown(textView As IAvalonTextView, args As KeyEventArgs)
        End Sub

        Public Overridable Sub KeyUp(textView As IAvalonTextView, args As KeyEventArgs)
        End Sub

        Public Overridable Sub PreviewKeyDown(textView As IAvalonTextView, args As KeyEventArgs)
        End Sub

        Public Overridable Sub PreviewKeyUp(textView As IAvalonTextView, args As KeyEventArgs)
        End Sub

        Public Overridable Sub TextInputStart(textView As IAvalonTextView, args As TextCompositionEventArgs)
        End Sub

        Public Overridable Sub TextInput(textView As IAvalonTextView, args As TextCompositionEventArgs)
        End Sub

        Public Overridable Sub TextInputUpdate(textView As IAvalonTextView, args As TextCompositionEventArgs)
        End Sub

        Public Overridable Sub PreviewTextInputStart(textView As IAvalonTextView, args As TextCompositionEventArgs)
        End Sub

        Public Overridable Sub PreviewTextInput(textView As IAvalonTextView, args As TextCompositionEventArgs)
        End Sub

        Public Overridable Sub PreviewTextInputUpdate(textView As IAvalonTextView, args As TextCompositionEventArgs)
        End Sub

    End Class
End Namespace
