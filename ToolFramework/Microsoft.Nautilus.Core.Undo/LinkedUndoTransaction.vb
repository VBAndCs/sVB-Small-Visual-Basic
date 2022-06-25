
Namespace Microsoft.Nautilus.Core.Undo
    Public MustInherit Class LinkedUndoTransaction
        Implements IDisposable

        Public MustOverride Property Description As String

        Public MustOverride Property CompletionTime As DateTime

        Public MustOverride ReadOnly Property State As UndoTransactionState

        Public MustOverride ReadOnly Property CanRedo As Boolean

        Public MustOverride ReadOnly Property CanUndo As Boolean

        Public MustOverride Property Policy As ILinkedUndoTransactionPolicy

        Public MustOverride Sub Complete()

        Public MustOverride Sub Cancel()

        Public MustOverride Sub Rollback()

        Public MustOverride Sub [Do]()

        Public MustOverride Sub Undo()

        Public Sub Dispose() Implements IDisposable.Dispose
            Dispose(disposing:=True)
            GC.SuppressFinalize(Me)
        End Sub

        Protected MustOverride Sub Dispose(disposing As Boolean)

    End Class
End Namespace
