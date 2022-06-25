Imports System.Collections.ObjectModel

Namespace Microsoft.Nautilus.Text.Editor
    Public Interface ITextLine
        ReadOnly Property IsDisposed As Boolean

        ReadOnly Property Baseline As Double

        ReadOnly Property Extent As Double

        ReadOnly Property OverhangAfter As Double

        ReadOnly Property OverhangLeading As Double

        ReadOnly Property OverhangTrailing As Double

        ReadOnly Property LineSpan As Span

        ReadOnly Property NewlineLength As Integer

        ReadOnly Property Left As Double

        ReadOnly Property Top As Double

        ReadOnly Property Height As Double

        ReadOnly Property Width As Double

        ReadOnly Property Bottom As Double

        ReadOnly Property Right As Double

        Function MoveCaretToLocation(horizontalPosition As Double) As ICaretPosition

        Function GetPositionFromXCoordinate(x As Double) As Integer?

        Function GetTextElementSpan(textElementIndex As Integer) As Span

        Function GetCharacterBounds(textBufferIndex As Integer) As TextBounds

        Function GetTextElementIndex(textBufferIndex As Integer) As Integer

        Function ContainsPosition(textBufferIndex As Integer) As Boolean

        Function GetTextBounds(span1 As Span) As ReadOnlyCollection(Of TextBounds)
    End Interface
End Namespace
