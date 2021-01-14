Enum ConnectionAction
    Create
    Delete
End Enum

Class ConnectionInfo
    Public SourceDiagram As FrameworkElement
    Public SourceConnector As ConnectorThumb
    Public TargetDiagram As FrameworkElement
    Public TargetConnector As ConnectorThumb

    Sub New(SourceDiagram As FrameworkElement,
                   SourceConnector As ConnectorThumb,
                   TargetDiagram As FrameworkElement,
                   TargetConnector As ConnectorThumb)

        Me.SourceDiagram = SourceDiagram
        Me.SourceConnector = SourceConnector

        Me.TargetDiagram = TargetDiagram
        Me.TargetConnector = TargetConnector
    End Sub

    Sub New(Connection As Connection)
        Me.New(Connection.SourceDiagram,
               Connection.SourceConnector,
               Connection.TargetDiagram,
               Connection.TargetConnector)
    End Sub
End Class

Class ConnectionState
    Inherits List(Of ConnectionInfo)
    Implements IRestore

    Public Action As ConnectionAction
    Public AfterRestoreAction As Action(Of List(Of Connection), ConnectionAction)

    Sub New(SourceDiagram As FrameworkElement,
                SourceConnector As ConnectorThumb,
                TargetDiagram As FrameworkElement,
                TargetConnector As ConnectorThumb,
                Action As ConnectionAction)

        Me.Add(New ConnectionInfo(SourceDiagram, SourceConnector, TargetDiagram, TargetConnector))
        Me.Action = Action
    End Sub

    Sub New(Action As ConnectionAction)
        Me.Action = Action
    End Sub

    Sub New(Connection As Connection, Action As ConnectionAction)
        Me.Add(New ConnectionInfo(Connection))
        Me.Action = Action
    End Sub

    Sub New(Action As ConnectionAction, AfterRestoreAction As Action(Of List(Of Connection), ConnectionAction))
        Me.Action = Action
        Me.AfterRestoreAction = AfterRestoreAction
    End Sub

    Sub New(Connection As Connection, Action As ConnectionAction, AfterRestoreAction As Action(Of List(Of Connection), ConnectionAction))
        Me.New(Connection, Action)
        Me.AfterRestoreAction = AfterRestoreAction
    End Sub

    Sub New(ConInfos As List(Of ConnectionInfo), Action As ConnectionAction, AfterRestoreAction As Action(Of List(Of Connection), ConnectionAction))
        Me.AddRange(ConInfos)
        Me.Action = Action
        Me.AfterRestoreAction = AfterRestoreAction
    End Sub

    Overloads Sub Add(Con As Connection)
        Me.Add(New ConnectionInfo(Con))
    End Sub

    Public Sub RestoreOldValue() Implements IRestore.RestoreOldValue
        If Me.Action = ConnectionAction.Delete Then
            Me.DeleteConnections()
        Else
            Me.CreateConnections()
        End If
    End Sub

    Public Sub RestoreNewValue() Implements IRestore.RestoreNewValue
        If Me.Action = ConnectionAction.Delete Then
            Me.CreateConnections()
        Else
            Me.DeleteConnections()
        End If
    End Sub

    Private Sub DeleteConnections()
        Dim ConList As New List(Of Connection)
        Dim Dsn = Helper.GetDesigner(Me(0).SourceDiagram)

        For Each ConInfo In Me

            Dim Conns = From C In Dsn.Connections(ConInfo.SourceDiagram)
                                  Where C.TargetDiagram Is ConInfo.TargetDiagram AndAlso
                                   C.SourceConnector Is ConInfo.SourceConnector

            If Conns.Count > 0 Then
                Dim Con = Conns.First
                Con.Delete()
                ConList.Add(Con)
            End If
        Next
        Dsn.Focus()
        If AfterRestoreAction IsNot Nothing Then AfterRestoreAction(ConList, Me.Action)

    End Sub

    Private Sub CreateConnections()
        Dim ConList As New List(Of Connection)
        Dim Dsn = Helper.GetDesigner(Me(0).SourceDiagram)

        For Each ConInfo In Me
            ConInfo.SourceConnector.Restore()
            Dsn.ConnectionSourceDiagram = ConInfo.SourceDiagram
            Dsn.SourceConnector = ConInfo.SourceConnector
            Dsn.SourceConnector.StartDrawConnection()

            ConInfo.TargetConnector.Restore()
            Dsn.ConnectionTargetDiagram = ConInfo.TargetDiagram
            Dsn.TargetConnector = ConInfo.TargetConnector
            Dim Con = ConInfo.SourceConnector.EndDrawConnection(True)
            Con.IsSelected = True
            ConList.Add(Con)
        Next
        If AfterRestoreAction IsNot Nothing Then AfterRestoreAction(ConList, Me.Action)

    End Sub

End Class