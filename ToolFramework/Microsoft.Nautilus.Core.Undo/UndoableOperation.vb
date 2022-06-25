Namespace Microsoft.Nautilus.Core.Undo
    Public Delegate Sub UndoableOperation(Of P0)(history As UndoHistory, p01 As P0)
    Public Delegate Sub UndoableOperation(Of P0, P1)(history As UndoHistory, p01 As P0, p11 As P1)
    Public Delegate Sub UndoableOperation(Of P0, P1, P2)(history As UndoHistory, p01 As P0, p11 As P1, p21 As P2)
    Public Delegate Sub UndoableOperation(Of P0, P1, P2, P3)(history As UndoHistory, p01 As P0, p11 As P1, p21 As P2, p31 As P3)
    Public Delegate Sub UndoableOperation(Of P0, P1, P2, P3, P4)(history As UndoHistory, p01 As P0, p11 As P1, p21 As P2, p31 As P3, p41 As P4)
End Namespace
