
Namespace Microsoft.Nautilus.Text
    Public Interface ITextBuffer
        Inherits IPropertyOwner
        Property ContentType As String

        ReadOnly Property CurrentSnapshot As ITextSnapshot

        ReadOnly Property ReadOnlyRegionManager As ReadOnlyRegionManager

        Event Changed As EventHandler(Of TextChangedEventArgs)

        Event ContentTypeChanged As EventHandler

        Function CreateEdit(sourceToken As Object) As ITextEdit

        Function CreateEdit() As ITextEdit

        Function Insert(position As Integer, text As String) As ITextSnapshot

        Function Delete(deleteSpan As Span) As ITextSnapshot

        Function Replace(replaceSpan As Span, replaceWith As String) As ITextSnapshot
    End Interface
End Namespace
