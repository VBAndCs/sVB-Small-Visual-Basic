Public Class Connection
    Inherits DependencyObject

    Dim SelectionPath As Path
    Friend ConnectionPath As Path
    Friend SourceConnector As ConnectorThumb
    Friend TargetConnector As ConnectorThumb
    Dim Dsn As Designer

    Dim _StraightLine As Boolean
    Public Property StraightLine As Boolean
        Get
            Return _StraightLine
        End Get
        Set(value As Boolean)
            _StraightLine = value
            UpdatePath()
        End Set
    End Property

    Friend Sub New(SourceConnector As ConnectorThumb, TargetConnector As ConnectorThumb, ConnectionPath As Path)
        SelectionPath = New Path
        SelectionPath.Visibility = Visibility.Collapsed
        SelectionPath.StrokeThickness = 1
        SelectionPath.Stroke = Brushes.Gray

        Me.SourceConnector = SourceConnector
        Me.TargetConnector = TargetConnector
        Me.ConnectionPath = ConnectionPath
        Connection.SetConnection(ConnectionPath, Me)
        Me.ConnectionPath.Focusable = True
        Me.ConnectionPath.FocusVisualStyle = Nothing

        Dsn = Helper.GetDesigner(SourceConnector)
        Dsn.ConnectionCanvas.Children.Add(SelectionPath)
        AddHandler Helper.GetDiagramPanel(SourceConnector).ConnectorsPositionChangd, AddressOf UpdatePath
        AddHandler Helper.GetDiagramPanel(TargetConnector).ConnectorsPositionChangd, AddressOf UpdatePath
        AddHandler ConnectionPath.MouseLeftButtonDown, AddressOf ConnectionPath_MouseLeftButtonDown
        AddHandler ConnectionPath.GotKeyboardFocus, AddressOf ConnectionPath_GotKeyboardFocus
        AddHandler ConnectionPath.LostKeyboardFocus, AddressOf ConnectionPath_LostKeyboardFocus
        AddHandler ConnectionPath.KeyDown, AddressOf ConnectionPath_KeyDown

        UpdatePath()
    End Sub

    Public Shared Function GetConnection(ByVal element As DependencyObject) As Connection
        If element Is Nothing Then
            Throw New ArgumentNullException("element")
        End If

        Return element.GetValue(ConnectionProperty)
    End Function

    Public Shared Sub SetConnection(ByVal element As DependencyObject, ByVal value As Connection)
        If element Is Nothing Then
            Throw New ArgumentNullException("element")
        End If

        element.SetValue(ConnectionProperty, value)
    End Sub

    Public Shared ReadOnly ConnectionProperty As  _
                           DependencyProperty = DependencyProperty.RegisterAttached("Connection", _
                           GetType(Connection), GetType(Connection))


    Private Function DrawArrow(point As Point, RotateAngle As Double) As PathGeometry
        Dim figure As New PathFigure()
        figure.StartPoint = point
        Dim X = point.X
        Dim Y = point.Y
        figure.Segments.Add(New PolyLineSegment({New Point(X, Y - 5), New Point(X + 10, Y), New Point(X, Y + 5), New Point(X, Y)}, True))
        figure.IsFilled = True
        Dim Geometry As New PathGeometry({figure})
        Geometry.Transform = New RotateTransform(RotateAngle, X, Y)
        Return Geometry
    End Function

    Public Sub Delete()

        Dim Pnl = Helper.GetDiagramPanel(SourceConnector)

        RemoveHandler Pnl.ConnectorsPositionChangd, AddressOf UpdatePath
        RemoveHandler Helper.GetDiagramPanel(TargetConnector).ConnectorsPositionChangd, AddressOf UpdatePath
        RemoveHandler ConnectionPath.LostKeyboardFocus, AddressOf ConnectionPath_LostKeyboardFocus

        'If Dsn.ConnectionCanvas.Children.Contains(ConnectionPath) Then
        Dsn.ConnectionCanvas.Children.Remove(ConnectionPath)
        Dsn.ConnectionCanvas.Children.Remove(SelectionPath)
        Dsn.GetConnectionsFromSource(Helper.GetDiagram(Pnl)).Remove(Me)
        SourceConnector.Delete()
        TargetConnector.Delete()
        'End If

        SelectionPath = Nothing
        SourceConnector = Nothing
        TargetConnector = Nothing
        ConnectionPath = Nothing
    End Sub

    Public ReadOnly Property SourceDiagram() As UIElement
        Get
            Dim Pnl = Helper.GetDiagramPanel(SourceConnector)
            Return Helper.GetDiagram(Pnl)
        End Get
    End Property

    Public ReadOnly Property TargetDiagram() As UIElement
        Get
            Dim Pnl = Helper.GetDiagramPanel(TargetConnector)
            Return Helper.GetDiagram(Pnl)
        End Get
    End Property

    Dim PathPoints As List(Of Point)

    Friend Sub UpdatePath()
        Dim geometry As New PathGeometry()
        Dim SrcInfo As New ConnectorInfo(SourceConnector)
        Dim TrgInfo As New ConnectorInfo(TargetConnector)
        PathPoints = PathFinder.GetConnectionLine(SrcInfo, TrgInfo, True, _StraightLine)
        If PathPoints.Count > 0 Then
            Dim figure As New PathFigure()
            figure.StartPoint = PathPoints(0)
            figure.Segments.Add(New PolyLineSegment(PathPoints, True))
            geometry.Figures.Add(figure)
            Dim Gg As New GeometryGroup
            Gg.Children.Add(geometry)
            Dim Angle As Double
            If _StraightLine Then
                Dim Tan = (PathPoints(1).Y - PathPoints(0).Y) / (PathPoints(1).X - PathPoints(0).X)
                Dim Arad = Math.Atan(Tan)
                Dim A2rad = Arad + 45 * Math.PI / 180
                Dim L = Math.Sqrt(50)
                Dim dx = Math.Cos(A2rad) * L
                Dim dy = Math.Sin(A2rad) * L

                SelectionPath.Margin = New Thickness(dx, dy, 0, 0)
                Dim xSign = If(PathPoints(1).X < PathPoints(0).X, 1, -1)
                Dim ySign = If(PathPoints(1).Y < PathPoints(0).Y, 1, -1)
                Angle = Math.Abs(Arad) * 180 / Math.PI
                If xSign = -1 Then
                    If ySign = 1 Then Angle = -Angle
                ElseIf ySign = -1 Then
                    Angle = 180 - Angle
                Else
                    Angle = 180 + Angle
                End If
            Else
                Angle = Me.TargetConnector.Orientation
                SelectionPath.Margin = New Thickness(5, 5, 0, 0)
            End If
            Gg.Children.Add(DrawArrow(PathPoints(PathPoints.Count - 1), Angle))
            Me.ConnectionPath.Data = Gg
            SelectionPath.Data = Gg
        End If
    End Sub

    Private Sub ConnectionPath_MouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs)
        Me.Focus()
        e.Handled = True
    End Sub

    Private Sub ConnectionPath_LostKeyboardFocus(sender As Object, e As KeyboardFocusChangedEventArgs)
        Me.Deselect(e.NewFocus)
    End Sub

    Private Sub ChangeSelection()
        If Keyboard.Modifiers = ModifierKeys.Control Then
            IsSelected = Not IsSelected
        ElseIf Not Me.IsSelected Then
            Dsn.SelectedIndex = -1
            Connection.DeselectAll(Dsn)
            Me.IsSelected = True
        End If
    End Sub

    Private Sub ConnectionPath_GotKeyboardFocus(sender As Object, e As KeyboardFocusChangedEventArgs)
        If e.OldFocus Is Me.ConnectionPath AndAlso Keyboard.Modifiers <> ModifierKeys.Control Then Return
        ChangeSelection()
        ConnectionPath.StrokeThickness = 2
    End Sub

    Private Sub ConnectionPath_KeyDown(sender As Object, e As KeyEventArgs)
        If e.Key = Key.Delete Then
            Dsn.RemoveSelectedItems()
        End If
    End Sub

    Property IsSelected As Boolean
        Get
            Return (SelectionPath.Visibility = Visibility.Visible)
        End Get
        Set(value As Boolean)
            SelectionPath.Visibility = If(value, Visibility.Visible, Visibility.Collapsed)
        End Set
    End Property


    Public Shared Sub SelectAll(Dsn As Designer)
        ChangeAllSellection(Dsn, True)
    End Sub

    Public Shared Sub DeselectAll(Dsn As Designer)
        ChangeAllSellection(Dsn, False)
    End Sub

    Private Shared Sub ChangeAllSellection(Dsn As Designer, NewState As Boolean)
        For Each ConList In Dsn.Connections.Values
            For Each Con In ConList
                Con.IsSelected = NewState
            Next
        Next
    End Sub

    Private Shared HitTestConnection As New List(Of Connection)

    Public Shared Sub SelectIntersection(Dsn As Designer, R As Rect)
        HitTestConnection.Clear()
        VisualTreeHelper.HitTest(Dsn.ConnectionCanvas, Nothing, AddressOf GetConnectionsHitTestResult, New GeometryHitTestParameters(New RectangleGeometry(R)))

        For Each Con As Connection In HitTestConnection
            Con.IsSelected = True
        Next

    End Sub

    Private Shared Function GetConnectionsHitTestResult(result As GeometryHitTestResult) As HitTestResultBehavior
        If Not result.IntersectionDetail = IntersectionDetail.Empty Then
            Dim Con = Connection.GetConnection(result.VisualHit)
            If Con IsNot Nothing Then HitTestConnection.Add(Con)
        End If
        Return HitTestResultBehavior.Continue
    End Function


    Public Sub Deselect(NewFocus As FrameworkElement)
        If Not ((TypeOf NewFocus Is Path OrElse TypeOf NewFocus Is ListBoxItem) AndAlso Keyboard.Modifiers = ModifierKeys.Control) Then
            Dim Con As Connection = Nothing
            Dim Itm As ListBoxItem = Nothing
            If NewFocus IsNot Nothing Then
                Con = TryCast(NewFocus.Tag, Connection)
                Itm = TryCast(NewFocus, ListBoxItem)
            End If
            If Not ((Con IsNot Nothing AndAlso Con.IsSelected) OrElse (Itm IsNot Nothing AndAlso Itm.IsSelected)) Then DeselectAll(Dsn)
        End If
        ConnectionPath.StrokeThickness = 1
    End Sub

    Public Sub Focus()
        If Keyboard.FocusedElement Is Me.ConnectionPath Then
            ChangeSelection()
        Else
            ConnectionPath.Focus()
        End If
    End Sub



End Class
