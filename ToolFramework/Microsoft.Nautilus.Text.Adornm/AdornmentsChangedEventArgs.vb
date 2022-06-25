
Namespace Microsoft.Nautilus.Text.AdornmentSystem
    Public Class AdornmentsChangedEventArgs
        Inherits EventArgs

        Public ReadOnly Property ChangeSpan As ITextSpan

        Public Sub New(changeSpan As ITextSpan)
            _ChangeSpan = changeSpan
        End Sub
    End Class
End Namespace
