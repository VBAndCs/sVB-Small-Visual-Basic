Class Stack(Of T)
    Inherits LinkedList(Of T)

    Friend Sub Push(item As T)
        MyBase.AddLast(item)
    End Sub

    Friend Function Peek() As T
        Return MyBase.Last.Value
    End Function

    Friend Function Pop() As T
        Pop = MyBase.Last.Value
        MyBase.RemoveLast()
    End Function

End Class

Public Class UndoRedoStack(Of T)
    Dim pUndo As New Stack(Of T)
    Dim pRedo As New Stack(Of T)
    Property LastChange As T

    Public MaxUndos As Integer = 100

    Public Event UndoRedoStateChanged(CanUndo As Boolean, CanRedo As Boolean)

    Sub New()

    End Sub

    Sub New(MaxUndos As Integer)
        Me.MaxUndos = MaxUndos
    End Sub

    Sub New(MaxUndos As Integer, InitialState As T)
        Me.MaxUndos = MaxUndos
        pUndo.Push(InitialState)
    End Sub

    ReadOnly Property CanUndo
        Get
            Return pUndo.Count > 0
        End Get
    End Property

    ReadOnly Property CanRedo
        Get
            Return pRedo.Count > 0
        End Get
    End Property

    Function Undo() As T
        pRedo.Push(pUndo.Peek)
        Undo = pUndo.Pop
        RaiseEvent UndoRedoStateChanged(pUndo.Count > 0, pRedo.Count > 0)
    End Function

    Function Redo() As T
        Dim R As T = pRedo.Pop
        pUndo.Push(R)
        RaiseEvent UndoRedoStateChanged(pUndo.Count > 0, pRedo.Count > 0)
        Return R
    End Function

    Sub ReportChanges(State As T, Optional HoldAsLastChange As Boolean = True)
        pRedo.Clear()
        pUndo.Push(State)
        If pUndo.Count > MaxUndos + 1 Then
            pUndo.RemoveFirst()
        End If
        RaiseEvent UndoRedoStateChanged(pUndo.Count > 0, pRedo.Count > 0)
        If HoldAsLastChange Then _LastChange = State
    End Sub

    Sub Clear()
        pRedo.Clear()
        pUndo.Clear()
        RaiseEvent UndoRedoStateChanged(False, False)
        _LastChange = Nothing
    End Sub


    Function Peek() As T
        If pUndo.Count = 0 Then Return Nothing
        Return pUndo.Peek
    End Function
End Class
