
Namespace Microsoft.Nautilus.Text
    Public Interface ITextEdit
        Inherits IDisposable
        ReadOnly Property Snapshot As ITextSnapshot

        Function CanInsert(position As Integer) As Boolean

        Function CanDeleteOrReplace(span1 As Span) As Boolean

        Function Insert(position As Integer, text As String) As Boolean

        Function Insert(position As Integer, characterBuffer As Char(), startIndex As Integer, length As Integer) As Boolean

        Function Delete(deleteSpan As Span) As Boolean

        Function Delete(startPosition As Integer, charsToDelete As Integer) As Boolean

        Function Replace(replaceSpan As Span, replaceWith As String) As Boolean

        Function Replace(startPosition As Integer, charsToReplace As Integer, replaceWith As String) As Boolean

        Function Apply() As ITextSnapshot

        Sub Cancel()
    End Interface
End Namespace
