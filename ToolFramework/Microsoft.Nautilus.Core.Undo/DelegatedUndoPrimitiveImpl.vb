Imports System.Collections.Generic

Namespace Microsoft.Nautilus.Core.Undo
    Friend Class DelegatedUndoPrimitiveImpl
        Implements IUndoPrimitive

        Private redoOperations As Stack(Of UndoableOperationCurried)
        Private undoOperations As Stack(Of UndoableOperationCurried)
        Private _parent As UndoTransactionImpl
        Private ReadOnly history As UndoHistoryImpl

        Public Property State As DelegatedUndoPrimitiveState

        Public ReadOnly Property CanRedo As Boolean Implements IUndoPrimitive.CanRedo
            Get
                Return redoOperations.Count > 0
            End Get
        End Property

        Public ReadOnly Property CanUndo As Boolean Implements IUndoPrimitive.CanUndo
            Get
                Return undoOperations.Count > 0
            End Get
        End Property

        Public Property Parent As UndoTransaction Implements IUndoPrimitive.Parent
            Get
                Return _parent
            End Get

            Set(value As UndoTransaction)
                _parent = TryCast(value, UndoTransactionImpl)
            End Set
        End Property

        Public ReadOnly Property MergeWithPreviousOnly As Boolean Implements IUndoPrimitive.MergeWithPreviousOnly
            Get
                Return True
            End Get
        End Property

        Public Sub New(history As UndoHistoryImpl, parent As UndoTransactionImpl, operationCurried As UndoableOperationCurried)
            redoOperations = New Stack(Of UndoableOperationCurried)
            undoOperations = New Stack(Of UndoableOperationCurried)
            Me._parent = parent
            Me.history = history
            _State = DelegatedUndoPrimitiveState.Inactive
            undoOperations.Push(operationCurried)
        End Sub

        Public Sub Undo() Implements IUndoPrimitive.Undo
            Using New CatchOperationsFromHistoryForDelegatedPrimitive(history, Me, DelegatedUndoPrimitiveState.Undoing)
                While undoOperations.Count > 0
                    undoOperations.Pop()()
                End While
            End Using
        End Sub

        Public Sub [Do]() Implements IUndoPrimitive.Do
            Using New CatchOperationsFromHistoryForDelegatedPrimitive(history, Me, DelegatedUndoPrimitiveState.Redoing)
                While redoOperations.Count > 0
                    redoOperations.Pop()()
                End While
            End Using
        End Sub

        Public Function GetEstimatedSize() As Long Implements IUndoPrimitive.GetEstimatedSize
            Return 0L
        End Function

        Public Sub AddOperation(operation As UndoableOperationCurried)
            If _State = DelegatedUndoPrimitiveState.Redoing Then
                undoOperations.Push(operation)
                Return
            End If

            If _State = DelegatedUndoPrimitiveState.Undoing Then
                redoOperations.Push(operation)
                Return
            End If
            Throw New InvalidOperationException("It is only possible to add a new operation to the DelegatedUndoPrimitive when the _State is Undoing or Redoing.")
        End Sub

        Public Function CanMerge(primitive As IUndoPrimitive) As Boolean Implements IUndoPrimitive.CanMerge
            Return False
        End Function

        Public Function Merge(primitive As IUndoPrimitive) As IUndoPrimitive Implements IUndoPrimitive.Merge
            Throw New InvalidOperationException("The DelegatedUndoPrimitive cannot Merge.")
        End Function
    End Class
End Namespace
