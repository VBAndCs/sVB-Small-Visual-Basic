Public Interface IRestore
    Sub RestoreOldValues()
    Sub RestoreNewValues()
End Interface

Public Class UndoRedoUnit
    Inherits List(Of IRestore)

    Sub New()

    End Sub

    Sub New(State As IRestore)
        Me.Add(State)
    End Sub

    Sub RestoreOldState()
        For Each State In Me
            State.RestoreOldValues()
        Next
    End Sub

    Sub RestoreNewState()
        For i = Me.Count - 1 To 0 Step -1
            Me(i).RestoreNewValues()
        Next
    End Sub

    Sub AddUnit(Unit As UndoRedoUnit)
        For Each U In Unit
            Me.Add(U)
        Next
    End Sub

End Class