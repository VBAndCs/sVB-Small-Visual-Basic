
Namespace Microsoft.Nautilus.Core.Undo
    Public MustInherit Class UndoPrimitive
        Implements IUndoPrimitive

        Public Overridable Property Parent As UndoTransaction Implements IUndoPrimitive.Parent

        Public Overridable ReadOnly Property CanRedo As Boolean = True Implements IUndoPrimitive.CanRedo

        Public Overridable ReadOnly Property CanUndo As Boolean = True Implements IUndoPrimitive.CanUndo

        Public Overridable ReadOnly Property MergeWithPreviousOnly As Boolean = True Implements IUndoPrimitive.MergeWithPreviousOnly

        Public Overridable Sub [Do]() Implements IUndoPrimitive.Do
            Toggle()
        End Sub

        Public Overridable Sub Undo() Implements IUndoPrimitive.Undo
            Toggle()
        End Sub

        Public Overridable Function GetEstimatedSize() As Long Implements IUndoPrimitive.GetEstimatedSize
            Return 0L
        End Function

        Protected Overridable Sub Toggle()
        End Sub

        Public Overridable Function CanMerge(other As IUndoPrimitive) As Boolean Implements IUndoPrimitive.CanMerge
            Return False
        End Function

        Public Overridable Function Merge(other As IUndoPrimitive) As IUndoPrimitive Implements IUndoPrimitive.Merge
            Throw New InvalidOperationException("The NullMergePolicy disallows merging of UndoPrimitives.")
        End Function

    End Class
End Namespace
