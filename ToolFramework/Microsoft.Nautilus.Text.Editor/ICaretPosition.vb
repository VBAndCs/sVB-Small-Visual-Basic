Namespace Microsoft.Nautilus.Text.Editor
    Public Interface ICaretPosition
        ReadOnly Property CharacterIndex As Integer

        ReadOnly Property Placement As CaretPlacement

        ReadOnly Property TextInsertionIndex As Integer
    End Interface
End Namespace
