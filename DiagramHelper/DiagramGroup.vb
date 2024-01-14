Friend Class DiagramGroup
    Private Shared Groups As New Dictionary(Of Long, DiagramGroup)

    Friend Panels As New List(Of DiagramPanel)
    Dim SelectionBorder As Border
    Dim Canvas As Canvas
    Dim timeStamp As Long

    Shared Sub Clear()
        Groups.Clear()
    End Sub

    Private Shared Function GetGroup(timeStamp As Long, Canvas As Canvas) As DiagramGroup
        If Groups.ContainsKey(timeStamp) Then
            Return Groups(timeStamp)
        Else
            Dim G As New DiagramGroup(Canvas)
            Groups.Add(timeStamp, G)
            Return G
        End If
    End Function

    Shared Sub Add(Pnl As DiagramPanel, timeStamp As Long)
        If timeStamp = 0 Then
            If Pnl.DiagramGroup IsNot Nothing Then Pnl.DiagramGroup.DoRemove(Pnl)
        Else
            Dim G = GetGroup(timeStamp, Pnl.Dsn.ConnectionCanvas)
            G.Add(Pnl)
            Pnl.ExitGroupChecked = True
            Pnl.GroupMenuItem.IsChecked = True
            Pnl.ExitGroupChecked = False
        End If
    End Sub

    Shared Sub RemovePanelOnly(Pnl As DiagramPanel)
        Dim G = Pnl.DiagramGroup
        If G Is Nothing Then Return
        G.DoRemove(Pnl, False)
    End Sub

    Private Shared Sub Remove(Pnl As DiagramPanel)
        Dim G = Pnl.DiagramGroup
        If G Is Nothing Then Return

        If G.Panels.Count < 3 Then
            Remove(G)
        Else
            G.DoRemove(Pnl)
        End If
    End Sub

    Private Shared Sub Remove(group As DiagramGroup)
        For Each g In Groups
            If g.Value Is group Then
                Groups.Remove(g.Key)
                Return
            End If
        Next
    End Sub

    Private Sub New(canvas As Canvas)
        SelectionBorder = New Border With {.IsHitTestVisible = False,
                      .Visibility = Visibility.Collapsed, .BorderBrush = Brushes.DarkRed,
                       .BorderThickness = New Thickness(1)}

        Me.Canvas = canvas
        canvas.Children.Add(SelectionBorder)
        timeStamp = Now.Ticks()
    End Sub

    Private Sub Add(Pnl As DiagramPanel)
        If Panels.Contains(Pnl) Then Return

        Panels.Add(Pnl)
        If Pnl.DiagramGroup IsNot Nothing Then Pnl.DiagramGroup.DoRemove(Pnl)
        Pnl.DiagramGroup = Me

        Dim rd = TryCast(Pnl.Diagram, RadioButton)
        If rd IsNot Nothing Then
            rd.GroupName = timeStamp.ToString()
        End If

        AddHandler Pnl.IsSelectedChanged, AddressOf Pnl_IsSelectedChanged
        UpdateSelection()
    End Sub

    Private Sub DoRemove(Pnl As DiagramPanel, Optional RemoveLastItem As Boolean = True)
        RemoveHandler Pnl.IsSelectedChanged, AddressOf Pnl_IsSelectedChanged

        Panels.Remove(Pnl)
        Dim rd = TryCast(Pnl.Diagram, RadioButton)
        If rd IsNot Nothing Then
            rd.GroupName = ""
        End If

        Pnl.ExitGroupChecked = True
        Pnl.GroupMenuItem.IsChecked = False
        Pnl.ExitGroupChecked = False

        Pnl.DiagramGroup = Nothing

        Select Case Panels.Count
            Case 0
                SelectionBorder.Visibility = Visibility.Collapsed
                Canvas.Children.Remove(SelectionBorder)
                Remove(Me)
            Case 1
                If RemoveLastItem Then
                    Panels(0).Diagram.ClearValue(Designer.GroupIDProperty)
                Else
                    UpdateSelection()
                End If
            Case Else
                UpdateSelection()
        End Select
    End Sub

    Private Sub Pnl_IsSelectedChanged(Pnl As DiagramPanel, NewValue As Boolean)
        For Each P In Panels
            If P Is Pnl Then Continue For
            P.ExitIsSelectedChanged = True
            P.IsSelected = NewValue
            P.ExitIsSelectedChanged = False
        Next

        If NewValue Then
            UpdateSelection()
            SelectionBorder.Visibility = Visibility.Visible
        Else
            SelectionBorder.Visibility = Visibility.Collapsed
        End If
    End Sub

    Friend Sub UpdateSelection()
        Dim MinX = Double.MaxValue
        Dim MinY = Double.MaxValue
        Dim MaxX = Double.MinValue
        Dim MaxY = Double.MinValue

        For Each Pnl In Panels
            Dim R As New Rect(0, 0, Pnl.ActualWidth, Pnl.ActualHeight)
            R = Pnl.TransformToVisual(Canvas).TransformBounds(R)
            MinX = Math.Min(MinX, R.Left)
            MaxX = Math.Max(MaxX, R.Left + R.Width)
            MinY = Math.Min(MinY, R.Top)
            MaxY = Math.Max(MaxY, R.Top + R.Height)
        Next

        Canvas.SetLeft(SelectionBorder, MinX - 30)
        Canvas.SetTop(SelectionBorder, MinY - 30)
        SelectionBorder.Width = MaxX - MinX + 60
        SelectionBorder.Height = MaxY - MinY + 60
    End Sub


    Sub [Select]()
        SelectionBorder.Visibility = Visibility.Visible
    End Sub

    ReadOnly Property Count As Integer
        Get
            Return Panels.Count
        End Get
    End Property

    Sub Ungroup()
        Dim UndoUnit As New UndoRedoUnit
        Do Until Panels.Count = 0
            Dim Diagram = Panels(0).Diagram
            Dim A As Action = AddressOf DiagramObject.Diagrams(Diagram).AfterRestoreAction
            Dim OldSate As New PropertyState(A, Diagram, Designer.GroupIDProperty)
            Diagram.ClearValue(Designer.GroupIDProperty)
            UndoUnit.Add(OldSate.SetNewValues)
        Loop
        Dim Dsn = Helper.GetDesigner(Canvas)
        Dsn.UndoStack.ReportChanges(UndoUnit)
    End Sub

End Class
