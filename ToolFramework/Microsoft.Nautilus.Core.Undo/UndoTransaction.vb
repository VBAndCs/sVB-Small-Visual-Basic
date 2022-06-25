Imports System.Collections.Generic

Namespace Microsoft.Nautilus.Core.Undo
    Public MustInherit Class UndoTransaction
        Implements IDisposable
        Public MustOverride Property Description As String

        Public MustOverride Property CompletionTime As DateTime

        Public MustOverride Property IsHidden As Boolean

        Public MustOverride ReadOnly Property State As UndoTransactionState

        Public MustOverride ReadOnly Property History As UndoHistory

        Public MustOverride ReadOnly Property UndoPrimitives As IList(Of IUndoPrimitive)

        Public MustOverride ReadOnly Property Parent As UndoTransaction

        Public MustOverride ReadOnly Property CanRedo As Boolean

        Public MustOverride ReadOnly Property CanUndo As Boolean

        Public MustOverride Property MergePolicy As IMergeUndoTransactionPolicy
        Public MustOverride Sub Complete()
        Public MustOverride Sub Cancel()
        Public MustOverride Sub Rollback()
        Public MustOverride Sub AddUndo(undo As IUndoPrimitive)
        Public MustOverride Sub [Do]()
        Public MustOverride Sub Undo()
        Public MustOverride Function GetEstimatedSize() As Long

        Public Sub Dispose() Implements IDisposable.Dispose
            Dispose(disposing:=True)
            GC.SuppressFinalize(Me)
        End Sub

        Protected MustOverride Sub Dispose(disposing As Boolean)
    End Class
End Namespace
