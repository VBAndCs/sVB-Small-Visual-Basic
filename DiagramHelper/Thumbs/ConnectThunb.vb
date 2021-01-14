Imports System.Windows.Controls.Primitives

Friend Class ConnectorThumb
    Inherits Thumb

    Dim Pnl As DiagramPanel
    Dim Item As ListBoxItem
    Dim Dsn As Designer
    Dim Diagram As FrameworkElement

    Dim ConnectionPath As Path
    Dim Connection As Connection
    Dim OldState As PropertyState
    Dim DontSetPos As Boolean


    Private Sub New()
        Me.Width = 10
        Me.Height = 10
        Me.Opacity = 0.7
        Me.BorderBrush = Brushes.Black
        Me.BorderThickness = New Thickness(1)
        Me.Cursor = Cursors.Hand
    End Sub

    Public Sub New(Diagram As FrameworkElement)
        Me.New()
        Me.Diagram = Diagram
        Pnl = Helper.GetDiagramPanel(Diagram)
        Item = Pnl.DesignerItem
        Dsn = Pnl.MyDesigner

        Dim P = Mouse.GetPosition(Diagram)
        Me.Position = P
        Pnl.ConnectorsGrid.Children.Add(Me)

    End Sub

    Public Sub New(Diagram As FrameworkElement, Pos As Point, AbsOrientation As ConnectorOrientation)
        Me.New()
        Me.Diagram = Diagram
        Pnl = Helper.GetDiagramPanel(Diagram)
        Item = Pnl.DesignerItem
        Dsn = Pnl.MyDesigner
        Me.Position = Pos
        Me.AbsOrientation = AbsOrientation
        Pnl.ConnectorsGrid.Children.Add(Me)
    End Sub

    Sub Restore()
        Pnl = Helper.GetDiagramPanel(Diagram)
        Item = Pnl.DesignerItem
        Dsn = Pnl.MyDesigner
        Pnl.ConnectorsGrid.Children.Add(Me)
    End Sub

    Friend Property Position As Point
        Get
            Return GetValue(PositionProperty)
        End Get

        Set(ByVal value As Point)
            SetValue(PositionProperty, value)
        End Set
    End Property

    Public Shared ReadOnly PositionProperty As DependencyProperty = _
                           DependencyProperty.Register("Position", _
                           GetType(Point), GetType(ConnectorThumb), New PropertyMetadata(AddressOf PositionChanged))

    Shared Sub PositionChanged(Con As ConnectorThumb, e As DependencyPropertyChangedEventArgs)
        If Not Con.DontSetPos Then Con.SetPos(e.NewValue)
        If Con.Connection IsNot Nothing Then
            Helper.UpdateControl(Con)
            Con.Connection.UpdatePath()
        End If

    End Sub

    ReadOnly Property Center As Point
        Get
            Return Me.TransformToAncestor(Dsn.DesignerCanvas).Transform(New Point(5, 5))
        End Get
    End Property

    Public ReadOnly Property DiagramRect As Rect
        Get
            Dim R = Diagram.RenderTransform.TransformBounds(New Rect(0, 0, Diagram.ActualWidth, Diagram.ActualHeight))
            R = Diagram.TransformToVisual(Dsn.DesignerCanvas).TransformBounds(R)
            Return R
        End Get
    End Property

    Function GetOffsetCenter(Offset As Double) As Point
        Dim offsetPoint As New Point()
        Dim Position = Me.Center
        Select Case Me.Orientation
            Case ConnectorOrientation.Left
                offsetPoint = New Point(Position.X - Offset, Position.Y)
            Case ConnectorOrientation.Top
                offsetPoint = New Point(Position.X, Position.Y - Offset)
            Case ConnectorOrientation.Right
                offsetPoint = New Point(Position.X + Offset, Position.Y)
            Case ConnectorOrientation.Bottom
                offsetPoint = New Point(Position.X, Position.Y + Offset)
            Case Else
        End Select

        Return offsetPoint
    End Function

    Private Sub ConnectorThumb_DragStarted(sender As Object, e As DragStartedEventArgs) Handles Me.DragStarted
        If Connection IsNot Nothing Then
            Connection.Focus()
            Dsn.ScrollIntoView(Diagram)
        End If

        OldState = New PropertyState(Me, PositionProperty)
    End Sub

    Private Sub ConnectThumb_DragDelta(sender As Object, e As DragDeltaEventArgs) Handles Me.DragDelta
        Dim P As New Point(Me.Position.X + e.HorizontalChange, Me.Position.Y + e.VerticalChange)
        Select Case Me.AbsOrientation
            Case ConnectorOrientation.Left
                P.X = Math.Max(P.X, Position.X + 1)
            Case ConnectorOrientation.Right
                P.X = Math.Min(P.X, Position.X - 1)
            Case ConnectorOrientation.Top
                P.Y = Math.Max(P.Y, Position.Y + 1)
            Case ConnectorOrientation.Bottom
                P.Y = Math.Min(P.Y, Position.Y - 1)
        End Select

        If Diagram.InputHitTest(P) Is Diagram Then Me.Position = P
    End Sub

    Private Sub ConnectorThumb_DragCompleted(sender As Object, e As DragCompletedEventArgs) Handles Me.DragCompleted
        If Me.Connection IsNot Nothing Then Me.Connection.UpdatePath()
        If OldState.HasChanges Then
            Dsn.UndoStack.ReportChanges(New UndoRedoUnit(OldState.SetNewValue))
        End If
    End Sub

    Friend Sub UpdateConnection()
        Dim position As Point = Mouse.GetPosition(Dsn.DesignerCanvas)
        Dim geometry As New PathGeometry()

        Dim linePoints As List(Of Point) = PathFinder.GetConnectionLine(New ConnectorInfo(Me), position, ConnectorOrientation.None)

        If linePoints.Count > 0 Then
            Dim figure As New PathFigure()
            figure.StartPoint = linePoints(0)
            linePoints.Remove(linePoints(0))
            figure.Segments.Add(New PolyLineSegment(linePoints, True))
            geometry.Figures.Add(figure)
        End If

        ConnectionPath.Data = geometry
    End Sub

    Friend AbsOrientation As ConnectorOrientation

    Friend ReadOnly Property Orientation As ConnectorOrientation
        Get
            Dim Rt = TryCast(Item.RenderTransform, RotateTransform)
            If Rt Is Nothing OrElse AbsOrientation = ConnectorOrientation.None Then Return AbsOrientation
            Dim A = AbsOrientation + Rt.Angle
            If A > 360 Then A -= 360
            If A > 315 OrElse A <= 45 Then Return ConnectorOrientation.Left
            If A > 45 AndAlso A <= 135 Then Return ConnectorOrientation.Top
            If A > 135 AndAlso A <= 225 Then Return ConnectorOrientation.Right
            If A > 225 AndAlso A <= 315 Then Return ConnectorOrientation.Bottom
            Return ConnectorOrientation.None
        End Get
    End Property


    Friend Sub StartDrawConnection()
        ConnectionPath = New Path
        ConnectionPath.IsHitTestVisible = False
        ConnectionPath.Cursor = Cursors.Hand
        ConnectionPath.Stroke = Brushes.Black

        ConnectionPath.StrokeThickness = 1
        Dsn.ConnectionCanvas.Children.Add(ConnectionPath)
        Pnl.AdjustConnectors()
    End Sub

    Sub CancelDrawConnection()
        Dsn.ConnectionCanvas.Children.Remove(ConnectionPath)
        Delete()
        Dsn.ConnectionMode = False
        Dsn.ConnectionSourceDiagram = Nothing
        Dsn.ConnectionTargetDiagram = Nothing
        Dsn.SourceConnector = Nothing
        Dsn.TargetConnector = Nothing
    End Sub

    Function EndDrawConnection(Optional IsUndo As Boolean = False) As Connection
        Helper.UpdateControl(Dsn.TargetConnector)
        ConnectionPath.IsHitTestVisible = True
        Dim Con As New Connection(Dsn.SourceConnector, Dsn.TargetConnector, ConnectionPath)
        Dsn.GetConnectionsFromSource(Dsn.ConnectionSourceDiagram).Add(Con)
        Me.Connection = Con
        Dsn.TargetConnector.Connection = Con
        Helper.GetDiagramPanel(Con.TargetConnector).AdjustConnectors()

        If Not IsUndo Then
            Dim OldState As New ConnectionState(Con, ConnectionAction.Delete)
            Dsn.UndoStack.ReportChanges(New UndoRedoUnit(OldState))
        End If

        Dsn.ConnectionMode = False
        Dsn.ConnectionSourceDiagram = Nothing
        Dsn.ConnectionTargetDiagram = Nothing
        Dsn.SourceConnector = Nothing
        Dsn.TargetConnector = Nothing

        Return Con
    End Function

    Private Sub SetPos(P As Point)
        Dim X1 = XStrokeBinarySeacrch(Diagram, P, New Point(0, P.Y))
        Dim X2 = XStrokeBinarySeacrch(Diagram, P, New Point(Diagram.ActualWidth, P.Y))

        Dim RightAlign As Boolean
        If P.X - X1 > X2 - P.X Then
            X1 = X2
            RightAlign = True
        End If

        Dim Y1 = YStrokeBinarySeacrch(Diagram, P, New Point(P.X, 0))
        Dim Y2 = YStrokeBinarySeacrch(Diagram, P, New Point(P.X, Diagram.ActualHeight))

        Dim BottomAlign As Boolean
        If P.Y - Y1 > Y2 - P.Y Then
            Y1 = Y2
            BottomAlign = True
        End If

        Dim MarginLeft = 0
        Dim MarginRight = 0
        Dim MarginTop = 0
        Dim MarginBottom = 0

        Me.HorizontalAlignment = Windows.HorizontalAlignment.Left
        Me.VerticalAlignment = Windows.VerticalAlignment.Top

        If Math.Abs(P.X - X1) < Math.Abs(P.Y - Y1) Then
            MarginTop = P.Y - Me.Height / 2
            P.X = X1
            If RightAlign Then
                Me.AbsOrientation = ConnectorOrientation.Right
                Me.HorizontalAlignment = Windows.HorizontalAlignment.Right
                MarginRight = Diagram.ActualWidth - P.X - Me.Width
            Else
                Me.AbsOrientation = ConnectorOrientation.Left
                MarginLeft = P.X - Me.Width
            End If
        Else
            MarginLeft = P.X - Me.Width / 2
            P.Y = Y1
            If BottomAlign Then
                Me.AbsOrientation = ConnectorOrientation.Bottom
                Me.VerticalAlignment = Windows.VerticalAlignment.Bottom
                MarginBottom = Diagram.ActualHeight - P.Y - Me.Height
            Else
                Me.AbsOrientation = ConnectorOrientation.Top
                MarginTop = P.Y - Me.Height
            End If
        End If

        Me.Margin = New Thickness(MarginLeft, MarginTop, MarginRight, MarginBottom)
        DontSetPos = True
        Me.Position = P
        DontSetPos = False
    End Sub

    Function XStrokeBinarySeacrch(Diagram As FrameworkElement, StartPoint As Point, EndPoint As Point) As Double
        If Math.Abs(EndPoint.X - StartPoint.X) < 0.1 Then Return StartPoint.X

        Dim MidPoint As New Point(StartPoint.X + (EndPoint.X - StartPoint.X) / 2, StartPoint.Y)
        If Diagram.InputHitTest(StartPoint) Is Diagram Then
            Return XStrokeBinarySeacrch(Diagram, MidPoint, EndPoint)
        Else
            Return XStrokeBinarySeacrch(Diagram, StartPoint, MidPoint)
        End If
    End Function

    Function YStrokeBinarySeacrch(Diagram As FrameworkElement, StartPoint As Point, EndPoint As Point) As Double
        If Math.Abs(EndPoint.Y - StartPoint.Y) < 0.1 Then Return StartPoint.Y

        Dim MidPoint As New Point(StartPoint.X, StartPoint.Y + (EndPoint.Y - StartPoint.Y) / 2)
        If Diagram.InputHitTest(StartPoint) Is Diagram Then
            Return YStrokeBinarySeacrch(Diagram, MidPoint, EndPoint)
        Else
            Return YStrokeBinarySeacrch(Diagram, StartPoint, MidPoint)
        End If
    End Function

    Sub Delete()
        Pnl.ConnectorsGrid.Children.Remove(Me)
        ConnectionPath = Nothing
        Connection = Nothing        
    End Sub

    Dim Cntx As ContextMenu
    Public Overrides Sub OnApplyTemplate()
        MyBase.OnApplyTemplate()
        Cntx = Me.GetTemplateChild("PART_ContextMenu")
        AddHandler Cntx.Opened, AddressOf ContextMenu_Opened
        Cntx.AddHandler(MenuItem.ClickEvent, New RoutedEventHandler(AddressOf MenuItem_Click))
    End Sub

    Private Sub ContextMenu_Opened(sender As Object, e As RoutedEventArgs)
        If Me.Connection.StraightLine Then
            Dim Rdo As RadioButton = Cntx.Items(1)
            Rdo.IsChecked = True
        Else
            Dim Rdo As RadioButton = Cntx.Items(0)
            Rdo.IsChecked = True
        End If
    End Sub

    Private Sub MenuItem_Click(sender As Object, e As RoutedEventArgs)
        Dim Mi As MenuItem = e.OriginalSource
        Select Case Cntx.ItemContainerGenerator.IndexFromContainer(Mi)
            Case 0 ' Right-Angle Line
                Me.Connection.StraightLine = False
            Case 1 'Straight Line
                Me.Connection.StraightLine = True
            Case 3 'Delete
                Me.Connection.Delete()
        End Select
    End Sub



End Class
