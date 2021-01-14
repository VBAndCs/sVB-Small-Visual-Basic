Class CollectionState
    Inherits Dictionary(Of Object, ValuePair(Of Integer))
    Implements IRestore

    Public AfterRestoreAction As Action(Of Object)
    Public Collection As IList

    Sub New(Collection As IList)
        Me.Collection = Collection
    End Sub

    Sub New(ParamArray Items() As Object)
        Me.Add(Items)
    End Sub

    Sub New(AfterRestoreAction As Action(Of Object), Collection As IList)
        Me.Collection = Collection
        Me.AfterRestoreAction = AfterRestoreAction
    End Sub

    Sub New(AfterRestoreAction As Action(Of Object))
        Me.AfterRestoreAction = AfterRestoreAction
    End Sub

    Sub New(Collection As IList, ParamArray Items() As Object)
        Me.Collection = Collection
        Me.Add(Items)
    End Sub

    Sub New(AfterRestoreAction As Action(Of Object), Collection As IList, ParamArray Items() As Object)
        Me.AfterRestoreAction = AfterRestoreAction
        Me.Collection = Collection
        Me.Add(Items)
    End Sub

    Overloads Sub Add(ParamArray Items() As Object)
        For Each C In Items
            Me.Add(C, New ValuePair(Of Integer)(Collection.IndexOf(C)))
        Next
    End Sub

    Sub RestoreOldValue() Implements IRestore.RestoreOldValue
        For i = Me.Count - 1 To 0 Step -1
            If Me.Values(i).OldValue = -1 Then
                Collection.Remove(Me.Keys(i))
            Else
                Collection.Insert(Me.Values(i).OldValue, Me.Keys(i))
            End If
        Next

        If AfterRestoreAction IsNot Nothing Then
            AfterRestoreAction(Me.Keys.ToList)
        End If

    End Sub

    Sub RestoreNewValue() Implements IRestore.RestoreNewValue
        For i = 0 To Me.Count - 1
            If Me.Values(i).NewValue = -1 Then
                Collection.Remove(Me.Keys(i))
            Else
                Collection.Insert(Me.Values(i).NewValue, Me.Keys(i))
            End If
        Next

        If AfterRestoreAction IsNot Nothing Then
            AfterRestoreAction(Me.Keys.ToList)
        End If

    End Sub


    Function SetNewValue() As CollectionState
        For Each Pair In Me
            Pair.Value.NewValue = Collection.IndexOf(Pair.Key)
        Next
        Return Me
    End Function

End Class
