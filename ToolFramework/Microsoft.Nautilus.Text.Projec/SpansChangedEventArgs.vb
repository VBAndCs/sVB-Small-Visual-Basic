Imports System.Collections.Generic

Namespace Microsoft.Nautilus.Text.Projection
    Public Class SpansChangedEventArgs
        Inherits EventArgs

        Public ReadOnly Property InsertedSpans As IList(Of ITextSpan)


        Public ReadOnly Property DeletedSpans As IList(Of ITextSpan)

        Public Sub New(insertedSpans As IList(Of ITextSpan), deletedSpans As IList(Of ITextSpan))
            If insertedSpans Is Nothing Then
                Throw New ArgumentNullException("insertedSpans")
            End If

            If deletedSpans Is Nothing Then
                Throw New ArgumentNullException("deletedSpans")
            End If

            _InsertedSpans = insertedSpans
            _DeletedSpans = deletedSpans
        End Sub
    End Class
End Namespace
