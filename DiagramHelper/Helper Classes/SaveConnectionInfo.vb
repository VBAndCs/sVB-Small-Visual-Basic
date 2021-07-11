Public Class SaveInfo
    Public Property Diagrams As ArrayList
    Public Property ConnectionInfos As ArrayList
    Public Property GridPen As Pen
    Public Property DesignerCanvasBrush As Brush
    Public Property PageSize As Size

    Sub New()

    End Sub

    Sub New(Dsn As Designer, Diagrams As IList, Connections As Dictionary(Of UIElement, List(Of Connection)))
        Me.GridPen = Dsn.GridPen
        Me.DesignerCanvasBrush = Dsn.DesignerCanvas.Background
        Me.PageSize = New Size(Dsn.PageWidth, Dsn.PageHeight)

        Me.Diagrams = New ArrayList()
        For Each Diagram In Diagrams
            Me.Diagrams.Add(Diagram)
        Next

        Me.ConnectionInfos = New ArrayList
        For Each ConList In Connections.Values
            For Each Con In ConList
                ConnectionInfos.Add(New SaveConnectionInfo(Dsn, Con))
            Next
        Next
    End Sub

    Sub LoadTo(Dsn As Designer)
        Dsn.PageWidth = Me.PageSize.Width
        Dsn.PageHeight = Me.PageSize.Height
        If Me.DesignerCanvasBrush IsNot Nothing Then Dsn.DesignerCanvas.Background = Me.DesignerCanvasBrush

        If GridPen IsNot Nothing Then
            Dsn.GridPen.Thickness = Me.GridPen.Thickness
            Dsn.GridPen.Brush = Me.GridPen.Brush
        End If


        For Each Diagram In Me.Diagrams
            Dsn.Items.Add(Diagram)
        Next

        Helper.UpdateControl(Dsn)

        For Each ConInfo As SaveConnectionInfo In ConnectionInfos
            ConInfo.CreateConnection(Dsn)
        Next
    End Sub

End Class

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

