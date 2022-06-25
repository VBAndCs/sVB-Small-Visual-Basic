
Namespace Microsoft.Nautilus.Core.Undo
    Friend Class CatchOperationsFromHistoryForDelegatedPrimitive
        Implements IDisposable

        Private history As UndoHistoryImpl
        Private primitive As DelegatedUndoPrimitiveImpl

        Public Sub New(history As UndoHistoryImpl, primitive As DelegatedUndoPrimitiveImpl, state As DelegatedUndoPrimitiveState)
            Me.history = history
            Me.primitive = primitive
            primitive.State = state
            history.ForwardToUndoOperation(primitive)
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            history.EndForwardToUndoOperation(primitive)
            primitive.State = DelegatedUndoPrimitiveState.Inactive
        End Sub
    End Class
End Namespace
