Namespace Microsoft.Nautilus.Text
    Public Interface ITextPoint
        ReadOnly Property TextBuffer As ITextBuffer

        ReadOnly Property TrackingMode As TrackingMode

        Function GetPoint(snapshot As ITextSnapshot) As SnapshotPoint

        Function GetPosition(snapshot As ITextSnapshot) As Integer

        Function GetCharacter(snapshot As ITextSnapshot) As Char
    End Interface
End Namespace
