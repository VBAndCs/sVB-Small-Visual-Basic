
Namespace Microsoft.Nautilus.Text.Editor
    Public Interface ITextCaret
        ReadOnly Property PreferredBounds As TextBounds

        ReadOnly Property Left As Double

        Property Width As Double

        ReadOnly Property Right As Double

        ReadOnly Property Top As Double

        ReadOnly Property Height As Double

        ReadOnly Property Bottom As Double

        ReadOnly Property Position As ICaretPosition
        Event PositionChanged As EventHandler(Of CaretPositionChangedEventArgs)

        Sub EnsureVisible()

        Function MoveTo(characterIndex As Integer) As ICaretPosition

        Function MoveTo(characterIndex As Integer, caretPlacement1 As CaretPlacement) As ICaretPosition

        Function MoveToNextCaretPosition() As ICaretPosition

        Function MoveToPreviousCaretPosition() As ICaretPosition

        Sub CapturePreferredBounds()

        Sub CapturePreferredHorizontalBounds()

        Sub CapturePreferredVerticalBounds()
    End Interface
End Namespace
