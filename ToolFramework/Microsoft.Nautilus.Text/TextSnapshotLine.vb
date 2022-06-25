Imports Microsoft.Nautilus.Text.StringRebuilder

Namespace Microsoft.Nautilus.Text
    Friend Class TextSnapshotLine
        Implements ITextSnapshotLine

        Private snapshot As ITextSnapshot
        Private lineSpan As LineSpan

        Public ReadOnly Property TextSnapshot As ITextSnapshot Implements ITextSnapshotLine.TextSnapshot
            Get
                Return snapshot
            End Get
        End Property

        Public ReadOnly Property LineNumber As Integer Implements ITextSnapshotLine.LineNumber
            Get
                Return lineSpan.LineNumber
            End Get
        End Property

        Public ReadOnly Property Start As Integer Implements ITextSnapshotLine.Start
            Get
                Return lineSpan.Start
            End Get
        End Property

        Public ReadOnly Property Length As Integer Implements ITextSnapshotLine.Length
            Get
                Return lineSpan.Length
            End Get
        End Property

        Public ReadOnly Property LengthIncludingLineBreak As Integer Implements ITextSnapshotLine.LengthIncludingLineBreak
            Get
                Return lineSpan.LengthIncludingLineBreak
            End Get
        End Property

        Public ReadOnly Property LineBreakLength As Integer Implements ITextSnapshotLine.LineBreakLength
            Get
                Return lineSpan.LineBreakLength
            End Get
        End Property

        Public ReadOnly Property [End] As Integer Implements ITextSnapshotLine.End
            Get
                Return lineSpan.End
            End Get
        End Property

        Public ReadOnly Property EndIncludingLineBreak As Integer Implements ITextSnapshotLine.EndIncludingLineBreak
            Get
                Return lineSpan.EndIncludingLineBreak
            End Get
        End Property

        Public ReadOnly Property Extent As Span Implements ITextSnapshotLine.Extent
            Get
                Return lineSpan.Extent
            End Get
        End Property

        Public ReadOnly Property ExtentIncludingLineBreak As Span Implements ITextSnapshotLine.ExtentIncludingLineBreak
            Get
                Return lineSpan.ExtentIncludingLineBreak
            End Get
        End Property

        Public Sub New(snapshot As ITextSnapshot, lineSpan As LineSpan)
            If snapshot Is Nothing Then
                Throw New ArgumentNullException("snapshot")
            End If

            If lineSpan.EndIncludingLineBreak > snapshot.Length Then
                Throw New ArgumentOutOfRangeException("lineSpan")
            End If

            Me.snapshot = snapshot
            Me.lineSpan = lineSpan
        End Sub

        Public Function GetText() As String Implements ITextSnapshotLine.GetText
            Return snapshot.GetText(lineSpan.Extent)
        End Function

        Public Function GetTextIncludingLineBreak() As String Implements ITextSnapshotLine.GetTextIncludingLineBreak
            Return snapshot.GetText(lineSpan.ExtentIncludingLineBreak)
        End Function

        Public Function GetLineBreakText() As String Implements ITextSnapshotLine.GetLineBreakText
            Return snapshot.GetText(New Span(lineSpan.End, lineSpan.LineBreakLength))
        End Function

        Public Function GetPositionOfNextNonWhiteSpaceCharacter(start1 As Integer) As Integer Implements ITextSnapshotLine.GetPositionOfNextNonWhiteSpaceCharacter
            Dim text As String = GetText()
            Dim length1 As Integer = text.Length
            For i As Integer = start1 To length1 - 1
                Dim c As Char = text(i)
                If Not Char.IsWhiteSpace(c) Then
                    Return i
                End If
            Next
            Return length1
        End Function
    End Class
End Namespace
