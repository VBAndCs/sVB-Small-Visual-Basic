Namespace Microsoft.Nautilus.Core.Undo
    Public Interface IUndoPrimitive
        Property Parent As UndoTransaction

        ReadOnly Property CanRedo As Boolean

        ReadOnly Property CanUndo As Boolean

        ReadOnly Property MergeWithPreviousOnly As Boolean

        Sub [Do]()

        Sub Undo()

        Function GetEstimatedSize() As Long

        Function CanMerge(other As IUndoPrimitive) As Boolean

        Function Merge(other As IUndoPrimitive) As IUndoPrimitive
    End Interface
End Namespace
