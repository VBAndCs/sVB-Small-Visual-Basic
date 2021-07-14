
Public Class SaveConnectionInfo
    Public Property SourceIndex As Integer
    Public Property SourcePos As Point
    Public Property SourceOrientation As ConnectorOrientation

    Public Property TargetIndex As Integer
    Public Property TargetPos As Point
    Public Property TargetOrientation As ConnectorOrientation

    Sub New()

    End Sub

    Sub New(Dsn As Designer, Connection As Connection)
        Me.SourceIndex = Dsn.Items.IndexOf(Connection.SourceDiagram)
        Me.SourcePos = Connection.SourceConnector.Position
        Me.SourceOrientation = Connection.SourceConnector.AbsOrientation

        Me.TargetIndex = Dsn.Items.IndexOf(Connection.TargetDiagram)
        Me.TargetPos = Connection.TargetConnector.Position
        Me.TargetOrientation = Connection.TargetConnector.AbsOrientation
    End Sub

    Function CreateConnection(Dsn As Designer) As Connection
        Dim SourceDiagram = Dsn.Items(SourceIndex)
        Dim SourceConnector As New ConnectorThumb(SourceDiagram, SourcePos, SourceOrientation)
        Dsn.ConnectionSourceDiagram = SourceDiagram
        Dsn.SourceConnector = SourceConnector
        Dsn.SourceConnector.StartDrawConnection()

        Dim TargetDiagram = Dsn.Items(TargetIndex)
        Dim TargetConnector As New ConnectorThumb(TargetDiagram, TargetPos, TargetOrientation)
        Dsn.ConnectionTargetDiagram = TargetDiagram
        Dsn.TargetConnector = TargetConnector
        Return SourceConnector.EndDrawConnection(True)

    End Function
End Class

