
Namespace Microsoft.Nautilus.Text.Classification
    Public Class ClassificationChangedEventArgs
        Inherits EventArgs

        Public ReadOnly Property ChangeSpan As ITextSpan

        Public Sub New(changeSpan As ITextSpan)
            If changeSpan Is Nothing Then
                Throw New ArgumentNullException("changeSpan")
            End If
            _ChangeSpan = changeSpan
        End Sub
    End Class
End Namespace
