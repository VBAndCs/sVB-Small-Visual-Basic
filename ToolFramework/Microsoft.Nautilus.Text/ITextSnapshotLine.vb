Namespace Microsoft.Nautilus.Text
    Public Interface ITextSnapshotLine
        ReadOnly Property TextSnapshot As ITextSnapshot

        ReadOnly Property Extent As Span

        ReadOnly Property ExtentIncludingLineBreak As Span

        ReadOnly Property LineNumber As Integer

        ReadOnly Property Start As Integer

        ReadOnly Property Length As Integer

        ReadOnly Property LengthIncludingLineBreak As Integer

        ReadOnly Property [End] As Integer

        ReadOnly Property EndIncludingLineBreak As Integer

        ReadOnly Property LineBreakLength As Integer

        Function GetText() As String

        Function GetTextIncludingLineBreak() As String

        Function GetLineBreakText() As String

        Function GetPositionOfNextNonWhiteSpaceCharacter(start1 As Integer) As Integer
    End Interface
End Namespace
