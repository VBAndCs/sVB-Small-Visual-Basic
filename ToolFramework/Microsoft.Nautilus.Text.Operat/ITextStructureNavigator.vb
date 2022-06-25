Namespace Microsoft.Nautilus.Text.Operations
    Public Interface ITextStructureNavigator
        ReadOnly Property ContentType As String

        Function GetExtentOfWord(currentPosition As SnapshotPoint) As TextExtent

        Function GetPositionOfStartOfLine(currentPosition As SnapshotPoint) As Integer

        Function GetPositionOfEndOfLine(currentPosition As SnapshotPoint) As Integer
    End Interface
End Namespace
