
Namespace Microsoft.Nautilus.Text.Editor
    Public Class TextViewLayoutChangedEventArgs
        Inherits EventArgs

        Public ReadOnly Property ChangeSpan As Span?

        Public ReadOnly Property OldSnapshot As ITextSnapshot

        Public ReadOnly Property NewSnapshot As ITextSnapshot

        Public Sub New(changeSpan As Span, oldSnapshot As ITextSnapshot, newSnapshot As ITextSnapshot)
            If oldSnapshot Is Nothing Then
                Throw New ArgumentNullException("oldSnapshot")
            End If

            If newSnapshot Is Nothing Then
                Throw New ArgumentNullException("newSnapshot")
            End If

            If oldSnapshot.TextBuffer IsNot newSnapshot.TextBuffer Then
                Throw New ArgumentException("newSnapshot")
            End If

            If oldSnapshot.Version > newSnapshot.Version Then
                Throw New ArgumentException("newSnapshot")
            End If

            _ChangeSpan = changeSpan
            _OldSnapshot = oldSnapshot
            _NewSnapshot = newSnapshot
        End Sub
    End Class
End Namespace
