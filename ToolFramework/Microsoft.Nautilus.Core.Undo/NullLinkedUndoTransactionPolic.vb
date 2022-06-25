Namespace Microsoft.Nautilus.Core.Undo
    Friend Class NullLinkedUndoTransactionPolicy
        Implements ILinkedUndoTransactionPolicy

        Public ReadOnly Property UndoRedoWhenLinksAreNotTopmostInHistory As Boolean = True Implements ILinkedUndoTransactionPolicy.UndoRedoWhenLinksAreNotTopmostInHistory

    End Class
End Namespace
