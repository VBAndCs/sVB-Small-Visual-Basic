Namespace Microsoft.Nautilus.Text
    Public Interface ITextChange
        ReadOnly Property Position As Integer

        ReadOnly Property Delta As Integer

        ReadOnly Property OldEnd As Integer

        ReadOnly Property NewEnd As Integer

        ReadOnly Property OldText As String

        ReadOnly Property NewText As String

        ReadOnly Property OldLength As Integer

        ReadOnly Property NewLength As Integer
    End Interface
End Namespace
