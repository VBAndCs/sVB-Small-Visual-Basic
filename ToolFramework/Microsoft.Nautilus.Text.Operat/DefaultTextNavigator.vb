
Namespace Microsoft.Nautilus.Text.Operations
    Friend Class DefaultTextNavigator
        Implements ITextStructureNavigator

        Private _textBuffer As ITextBuffer

        Public ReadOnly Property ContentType As String = "bin" Implements ITextStructureNavigator.ContentType

        Friend Sub New(textBuffer As ITextBuffer)
            _textBuffer = textBuffer
        End Sub

        Public Function GetPositionOfStartOfLine(currentPosition As SnapshotPoint) As Integer Implements ITextStructureNavigator.GetPositionOfStartOfLine
            If currentPosition.Snapshot.TextBuffer IsNot _textBuffer Then
                Throw New ArgumentException("currentPosition TextBuffer does not match to the current TextBuffer")
            End If

            Return currentPosition.GetContainingLine().Start
        End Function

        Public Function GetPositionOfEndOfLine(currentPosition As SnapshotPoint) As Integer Implements ITextStructureNavigator.GetPositionOfEndOfLine
            If currentPosition.Snapshot.TextBuffer IsNot _textBuffer Then
                Throw New ArgumentException("currentPosition TextBuffer does not match to the current TextBuffer")
            End If

            Return currentPosition.GetContainingLine().End
        End Function

        Public Function GetExtentOfWord(currentPosition As SnapshotPoint) As TextExtent Implements ITextStructureNavigator.GetExtentOfWord
            If currentPosition.Snapshot.TextBuffer IsNot _textBuffer Then
                Throw New ArgumentException("currentPosition TextBuffer does not match to the current TextBuffer")
            End If

            Dim pos = currentPosition.Position
            If currentPosition.Position >= currentPosition.Snapshot.Length - 1 Then
                Return New TextExtent(pos, currentPosition.Snapshot.Length - pos, isSignificant:=True)
            End If

            Return New TextExtent(pos, 1, isSignificant:=True)
        End Function
    End Class
End Namespace
