
Namespace Microsoft.Nautilus.Text.Editor
    Public Class CaretPositionChangedEventArgs
        Inherits EventArgs

        Public ReadOnly Property TextView As ITextView

        Public ReadOnly Property OldPosition As ICaretPosition

        Public ReadOnly Property NewPosition As ICaretPosition

        Public Sub New(textView As ITextView, oldPosition As ICaretPosition, newPosition As ICaretPosition)
            _TextView = textView
            _OldPosition = oldPosition
            _NewPosition = newPosition
        End Sub
    End Class
End Namespace
