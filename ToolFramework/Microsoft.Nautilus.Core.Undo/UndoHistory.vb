Imports System.Collections.Generic
Imports Microsoft.Nautilus.Core.Undo.DataExports
Imports System.Runtime.InteropServices

Namespace Microsoft.Nautilus.Core.Undo
    Public MustInherit Class UndoHistory

        Public MustOverride ReadOnly Property UndoStack As Stack(Of UndoTransaction)

        Public MustOverride ReadOnly Property RedoStack As Stack(Of UndoTransaction)

        Public MustOverride ReadOnly Property CanUndo As Boolean

        Public MustOverride ReadOnly Property CanRedo As Boolean

        Public MustOverride ReadOnly Property UndoDescription As String

        Public MustOverride ReadOnly Property RedoDescription As String

        Public MustOverride ReadOnly Property CurrentTransaction As UndoTransaction

        Public MustOverride ReadOnly Property State As UndoHistoryState

        Public Event UndoRedoHappened As EventHandler(Of UndoRedoEventArgs)

        Public MustOverride Function GetEstimatedSize() As Long

        Public MustOverride Function CreateTransaction(description As String) As UndoTransaction

        Public MustOverride Function CreateTransaction(description As String, isHidden As Boolean) As UndoTransaction

        Public MustOverride Sub Undo(count As Integer)

        Public MustOverride Sub Redo(count As Integer)

        Public MustOverride Sub [Do](undoPrimitive As IUndoPrimitive)

        Public MustOverride Sub AddUndo(operation As UndoableOperationVoid)

        Public MustOverride Sub AddUndo(Of P0)(operation As UndoableOperation(Of P0), p01 As P0)

        Public MustOverride Sub AddUndo(Of P0, P1)(operation As UndoableOperation(Of P0, P1), p01 As P0, p11 As P1)

        Public MustOverride Sub AddUndo(Of P0, P1, P2)(operation As UndoableOperation(Of P0, P1, P2), p01 As P0, p11 As P1, p21 As P2)

        Public MustOverride Sub AddUndo(Of P0, P1, P2, P3)(operation As UndoableOperation(Of P0, P1, P2, P3), p01 As P0, p11 As P1, p21 As P2, p31 As P3)

        Public MustOverride Sub AddUndo(Of P0, P1, P2, P3, P4)(operation As UndoableOperation(Of P0, P1, P2, P3, P4), p01 As P0, p11 As P1, p21 As P2, p31 As P3, p41 As P4)

        Public MustOverride Sub ReplaceMarker(marker As UndoTransactionMarker, transaction As UndoTransaction, _value As Object)

        Public MustOverride Function GetMarker(marker As UndoTransactionMarker, transaction As UndoTransaction) As Object

        Public MustOverride Sub ClearMarker(marker As UndoTransactionMarker)

        Public MustOverride Function FindMarker(marker As UndoTransactionMarker) As UndoTransaction

        Public MustOverride Function TryFindMarkerOnTop(marker As UndoTransactionMarker, <Out> ByRef _value As Object) As Boolean

        Public MustOverride Sub ReplaceMarkerOnTop(marker As UndoTransactionMarker, _value As Object)

        Protected Sub RaiseUndoRedoHappened(state As UndoHistoryState, transaction As UndoTransaction)
            RaiseEvent UndoRedoHappened(Me, New UndoRedoEventArgs(state, transaction))
        End Sub

    End Class
End Namespace
