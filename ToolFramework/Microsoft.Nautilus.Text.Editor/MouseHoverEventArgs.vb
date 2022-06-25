
Namespace Microsoft.Nautilus.Text.Editor
    Public Class MouseHoverEventArgs
        Inherits EventArgs

        Public ReadOnly Property View As ITextView

        Public ReadOnly Property Position As Integer

        Public Sub New(view As ITextView, position As Integer)
            If view Is Nothing Then
                Throw New ArgumentNullException("view")
            End If

            If position < 0 OrElse position >= view.TextSnapshot.Length Then
                Throw New ArgumentOutOfRangeException("position")
            End If

            _View = view
            _Position = position
        End Sub
    End Class
End Namespace
