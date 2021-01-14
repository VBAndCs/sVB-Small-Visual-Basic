Imports System
Imports System.Collections.Generic
Imports System.Windows
Imports System.Windows.Controls

Friend Structure ConnectorInfo
    Dim Diagram As FrameworkElement
    Dim Canv As Canvas

    Sub New(Connector As ConnectorThumb)
        Dim Pnl = Helper.GetDiagramPanel(Connector)
        Diagram = Helper.GetDiagram(Pnl)
        Dim Item = Pnl.DesignerItem
        Canv = Pnl.MyDesigner.DesignerCanvas
        Position = Connector.TransformToAncestor(Canv).Transform(New Point(5, 5))
        Me.Orientation = Connector.Orientation
    End Sub

    Public ReadOnly Property DiagramRect As Rect
        Get
            Dim R = Diagram.RenderTransform.TransformBounds(New Rect(0, 0, Diagram.ActualWidth, Diagram.ActualHeight))
            R = Diagram.TransformToVisual(Canv).TransformBounds(R)
            Return R
        End Get
    End Property

    Public Position As Point
    Public Orientation As ConnectorOrientation
End Structure

Public Enum ConnectorOrientation
    None = -1
    Left = 0
    Top = 90
    Right = 180
    Bottom = 270
End Enum

Friend Class PathFinder
    Private Const margin As Integer = 35

    Friend Shared Function GetConnectionLine(ByVal source As ConnectorInfo, ByVal Target As ConnectorInfo, ByVal showLastLine As Boolean, Optional StraitLine As Boolean = False) As List(Of Point)
        Dim linePoints As New List(Of Point)()

        If StraitLine Then
            linePoints.Add(source.Position)
            Dim m = Math.Abs((Target.Position.Y - source.Position.Y) / (Target.Position.X - source.Position.X))
            Dim dx = 18 / Math.Sqrt(1 + m ^ 2)
            Dim P = Target.Position
            Dim xSign = If(Target.Position.X < source.Position.X, 1, -1)
            Dim ySign = If(Target.Position.Y < source.Position.Y, 1, -1)
            P.Offset(xSign * dx, ySign * m * dx)
            linePoints.Add(P)
            Return linePoints
        End If

        Dim rectSource As Rect = GetRectWithMargin(source, margin)
        Dim rectTarget As Rect = GetRectWithMargin(Target, margin)

        Dim startPoint As Point = GetOffsetPoint(source, rectSource)
        Dim endPoint As Point = GetOffsetPoint(Target, rectTarget)


        linePoints.Add(startPoint)

        Dim currentPoint As Point = startPoint

        If (Not rectTarget.Contains(currentPoint)) AndAlso (Not rectSource.Contains(endPoint)) Then
            Do
                '#Region "source node"

                If IsPointVisible(currentPoint, endPoint, New Rect() {rectSource, rectTarget}) Then
                    linePoints.Add(endPoint)
                    currentPoint = endPoint
                    Exit Do
                End If

                Dim neighbour As Point = GetNearestVisibleNeighborTarget(currentPoint, endPoint, Target, rectSource, rectTarget)
                If Not Double.IsNaN(neighbour.X) Then
                    linePoints.Add(neighbour)
                    linePoints.Add(endPoint)
                    currentPoint = endPoint
                    Exit Do
                End If

                If currentPoint = startPoint Then
                    Dim flag As Boolean
                    Dim n As Point = GetNearestNeighborSource(source, endPoint, rectSource, rectTarget, flag)
                    linePoints.Add(n)
                    currentPoint = n

                    If Not IsRectVisible(currentPoint, rectTarget, New Rect() {rectSource}) Then
                        Dim n1, n2 As Point
                        GetOppositeCorners(source.Orientation, rectSource, n1, n2)
                        If flag Then
                            linePoints.Add(n1)
                            currentPoint = n1
                        Else
                            linePoints.Add(n2)
                            currentPoint = n2
                        End If
                        If Not IsRectVisible(currentPoint, rectTarget, New Rect() {rectSource}) Then
                            If flag Then
                                linePoints.Add(n2)
                                currentPoint = n2
                            Else
                                linePoints.Add(n1)
                                currentPoint = n1
                            End If
                        End If
                    End If
                    '#End Region

                    '#Region "Target node"

                Else ' from here on we jump to the Target node
                    Dim n1, n2 As Point ' neighbour corner
                    Dim s1, s2 As Point ' opposite corner
                    GetNeighborCorners(Target.Orientation, rectTarget, s1, s2)
                    GetOppositeCorners(Target.Orientation, rectTarget, n1, n2)

                    Dim n1Visible As Boolean = IsPointVisible(currentPoint, n1, New Rect() {rectSource, rectTarget})
                    Dim n2Visible As Boolean = IsPointVisible(currentPoint, n2, New Rect() {rectSource, rectTarget})

                    If n1Visible AndAlso n2Visible Then
                        If rectSource.Contains(n1) Then
                            linePoints.Add(n2)
                            If rectSource.Contains(s2) Then
                                linePoints.Add(n1)
                                linePoints.Add(s1)
                            Else
                                linePoints.Add(s2)
                            End If

                            linePoints.Add(endPoint)
                            currentPoint = endPoint
                            Exit Do
                        End If

                        If rectSource.Contains(n2) Then
                            linePoints.Add(n1)
                            If rectSource.Contains(s1) Then
                                linePoints.Add(n2)
                                linePoints.Add(s2)
                            Else
                                linePoints.Add(s1)
                            End If

                            linePoints.Add(endPoint)
                            currentPoint = endPoint
                            Exit Do
                        End If

                        If (Distance(n1, endPoint) <= Distance(n2, endPoint)) Then
                            linePoints.Add(n1)
                            If rectSource.Contains(s1) Then
                                linePoints.Add(n2)
                                linePoints.Add(s2)
                            Else
                                linePoints.Add(s1)
                            End If
                            linePoints.Add(endPoint)
                            currentPoint = endPoint
                            Exit Do
                        Else
                            linePoints.Add(n2)
                            If rectSource.Contains(s2) Then
                                linePoints.Add(n1)
                                linePoints.Add(s1)
                            Else
                                linePoints.Add(s2)
                            End If
                            linePoints.Add(endPoint)
                            currentPoint = endPoint
                            Exit Do
                        End If
                    ElseIf n1Visible Then
                        linePoints.Add(n1)
                        If rectSource.Contains(s1) Then
                            linePoints.Add(n2)
                            linePoints.Add(s2)
                        Else
                            linePoints.Add(s1)
                        End If
                        linePoints.Add(endPoint)
                        currentPoint = endPoint
                        Exit Do
                    Else
                        linePoints.Add(n2)
                        If rectSource.Contains(s2) Then
                            linePoints.Add(n1)
                            linePoints.Add(s1)
                        Else
                            linePoints.Add(s2)
                        End If
                        linePoints.Add(endPoint)
                        currentPoint = endPoint
                        Exit Do
                    End If
                End If
                '#End Region
            Loop
        Else
            linePoints.Add(endPoint)
        End If

        linePoints = OptimizeLinePoints(linePoints, New Rect() {rectSource, rectTarget}, source.Orientation, Target.Orientation)

        CheckPathEnd(source, Target, showLastLine, linePoints)
        Return linePoints
    End Function

    Friend Shared Function GetConnectionLine(ByVal source As ConnectorInfo, ByVal TargetPoint As Point, ByVal preferredOrientation As ConnectorOrientation) As List(Of Point)
        Dim linePoints As New List(Of Point)()
        Dim rectSource As Rect = GetRectWithMargin(source, 10)
        Dim startPoint As Point = GetOffsetPoint(source, rectSource)
        Dim endPoint As Point = TargetPoint

        linePoints.Add(startPoint)
        Dim currentPoint As Point = startPoint

        If Not rectSource.Contains(endPoint) Then
            Do
                If IsPointVisible(currentPoint, endPoint, New Rect() {rectSource}) Then
                    linePoints.Add(endPoint)
                    Exit Do
                End If

                Dim sideFlag As Boolean
                Dim n As Point = GetNearestNeighborSource(source, endPoint, rectSource, sideFlag)
                linePoints.Add(n)
                currentPoint = n

                If IsPointVisible(currentPoint, endPoint, New Rect() {rectSource}) Then
                    linePoints.Add(endPoint)
                    Exit Do
                Else
                    Dim n1, n2 As Point
                    GetOppositeCorners(source.Orientation, rectSource, n1, n2)
                    If sideFlag Then
                        linePoints.Add(n1)
                    Else
                        linePoints.Add(n2)
                    End If

                    linePoints.Add(endPoint)
                    Exit Do
                End If
            Loop
        Else
            linePoints.Add(endPoint)
        End If

        If preferredOrientation <> ConnectorOrientation.None Then
            linePoints = OptimizeLinePoints(linePoints, New Rect() {rectSource}, source.Orientation, preferredOrientation)
        Else
            linePoints = OptimizeLinePoints(linePoints, New Rect() {rectSource}, source.Orientation, GetOpositeOrientation(source.Orientation))
        End If

        Return linePoints
    End Function

    Private Shared Function OptimizeLinePoints(ByVal linePoints As List(Of Point), ByVal rectangles() As Rect, ByVal sourceOrientation As ConnectorOrientation, ByVal TargetOrientation As ConnectorOrientation) As List(Of Point)
        Dim points As New List(Of Point)()
        Dim cut As Integer = 0

        For i As Integer = 0 To linePoints.Count - 1
            If i >= cut Then
                For k As Integer = linePoints.Count - 1 To i + 1 Step -1
                    If IsPointVisible(linePoints(i), linePoints(k), rectangles) Then
                        cut = k
                        Exit For
                    End If
                Next k
                points.Add(linePoints(i))
            End If
        Next i

        '			#Region "Line"
        For j As Integer = 0 To points.Count - 2
            If points(j).X <> points(j + 1).X AndAlso points(j).Y <> points(j + 1).Y Then
                Dim orientationFrom As ConnectorOrientation
                Dim orientationTo As ConnectorOrientation

                ' orientation from point
                If j = 0 Then
                    orientationFrom = sourceOrientation
                Else
                    orientationFrom = GetOrientation(points(j), points(j - 1))
                End If

                ' orientation to pint 
                If j = points.Count - 2 Then
                    orientationTo = TargetOrientation
                Else
                    orientationTo = GetOrientation(points(j + 1), points(j + 2))
                End If


                If (orientationFrom = ConnectorOrientation.Left OrElse orientationFrom = ConnectorOrientation.Right) AndAlso (orientationTo = ConnectorOrientation.Left OrElse orientationTo = ConnectorOrientation.Right) Then
                    Dim centerX As Double = Math.Min(points(j).X, points(j + 1).X) + Math.Abs(points(j).X - points(j + 1).X) \ 2
                    points.Insert(j + 1, New Point(centerX, points(j).Y))
                    points.Insert(j + 2, New Point(centerX, points(j + 2).Y))
                    If points.Count - 1 > j + 3 Then
                        points.RemoveAt(j + 3)
                    End If
                    Return points
                End If

                If (orientationFrom = ConnectorOrientation.Top OrElse orientationFrom = ConnectorOrientation.Bottom) AndAlso (orientationTo = ConnectorOrientation.Top OrElse orientationTo = ConnectorOrientation.Bottom) Then
                    Dim centerY As Double = Math.Min(points(j).Y, points(j + 1).Y) + Math.Abs(points(j).Y - points(j + 1).Y) \ 2
                    points.Insert(j + 1, New Point(points(j).X, centerY))
                    points.Insert(j + 2, New Point(points(j + 2).X, centerY))
                    If points.Count - 1 > j + 3 Then
                        points.RemoveAt(j + 3)
                    End If
                    Return points
                End If

                If (orientationFrom = ConnectorOrientation.Left OrElse orientationFrom = ConnectorOrientation.Right) AndAlso (orientationTo = ConnectorOrientation.Top OrElse orientationTo = ConnectorOrientation.Bottom) Then
                    points.Insert(j + 1, New Point(points(j + 1).X, points(j).Y))
                    Return points
                End If

                If (orientationFrom = ConnectorOrientation.Top OrElse orientationFrom = ConnectorOrientation.Bottom) AndAlso (orientationTo = ConnectorOrientation.Left OrElse orientationTo = ConnectorOrientation.Right) Then
                    points.Insert(j + 1, New Point(points(j).X, points(j + 1).Y))
                    Return points
                End If
            End If
        Next j
        '			#End Region

        Return points
    End Function

    Private Shared Function GetOrientation(ByVal p1 As Point, ByVal p2 As Point) As ConnectorOrientation
        If p1.X = p2.X Then
            If p1.Y >= p2.Y Then
                Return ConnectorOrientation.Bottom
            Else
                Return ConnectorOrientation.Top
            End If
        ElseIf p1.Y = p2.Y Then
            If p1.X >= p2.X Then
                Return ConnectorOrientation.Right
            Else
                Return ConnectorOrientation.Left
            End If
        End If
        Return ConnectorOrientation.Left
    End Function

    Private Shared Function GetOrientation(ByVal sourceOrientation As ConnectorOrientation) As Orientation
        Select Case sourceOrientation
            Case ConnectorOrientation.Left
                Return Orientation.Horizontal
            Case ConnectorOrientation.Top
                Return Orientation.Vertical
            Case ConnectorOrientation.Right
                Return Orientation.Horizontal
            Case ConnectorOrientation.Bottom
                Return Orientation.Vertical
            Case Else
                Throw New Exception("Unknown ConnectorOrientation")
        End Select
    End Function

    Private Shared Function GetNearestNeighborSource(ByVal source As ConnectorInfo, ByVal endPoint As Point, ByVal rectSource As Rect, ByVal rectTarget As Rect, ByRef flag As Boolean) As Point
        Dim n1, n2 As Point ' neighbors
        GetNeighborCorners(source.Orientation, rectSource, n1, n2)

        If rectTarget.Contains(n1) Then
            flag = False
            Return n2
        End If

        If rectTarget.Contains(n2) Then
            flag = True
            Return n1
        End If

        If (Distance(n1, endPoint) <= Distance(n2, endPoint)) Then
            flag = True
            Return n1
        Else
            flag = False
            Return n2
        End If
    End Function

    Private Shared Function GetNearestNeighborSource(ByVal source As ConnectorInfo, ByVal endPoint As Point, ByVal rectSource As Rect, ByRef flag As Boolean) As Point
        Dim n1, n2 As Point ' neighbors
        GetNeighborCorners(source.Orientation, rectSource, n1, n2)

        If (Distance(n1, endPoint) <= Distance(n2, endPoint)) Then
            flag = True
            Return n1
        Else
            flag = False
            Return n2
        End If
    End Function

    Private Shared Function GetNearestVisibleNeighborTarget(ByVal currentPoint As Point, ByVal endPoint As Point, ByVal Target As ConnectorInfo, ByVal rectSource As Rect, ByVal rectTarget As Rect) As Point
        Dim s1, s2 As Point ' neighbors on Target side
        GetNeighborCorners(Target.Orientation, rectTarget, s1, s2)

        Dim flag1 As Boolean = IsPointVisible(currentPoint, s1, New Rect() {rectSource, rectTarget})
        Dim flag2 As Boolean = IsPointVisible(currentPoint, s2, New Rect() {rectSource, rectTarget})

        If flag1 Then ' s1 visible
            If flag2 Then ' s1 and s2 visible
                If rectTarget.Contains(s1) Then
                    Return s2
                End If

                If rectTarget.Contains(s2) Then
                    Return s1
                End If

                If (Distance(s1, endPoint) <= Distance(s2, endPoint)) Then
                    Return s1
                Else
                    Return s2
                End If

            Else
                Return s1
            End If
        Else ' s1 not visible
            If flag2 Then ' only s2 visible
                Return s2
            Else ' s1 and s2 not visible
                Return New Point(Double.NaN, Double.NaN)
            End If
        End If
    End Function

    Private Shared Function IsPointVisible(ByVal fromPoint As Point, ByVal targetPoint As Point, ByVal rectangles() As Rect) As Boolean
        For Each rect As Rect In rectangles
            If RectangleIntersectsLine(rect, fromPoint, targetPoint) Then
                Return False
            End If
        Next rect
        Return True
    End Function

    Private Shared Function IsRectVisible(ByVal fromPoint As Point, ByVal targetRect As Rect, ByVal rectangles() As Rect) As Boolean
        If IsPointVisible(fromPoint, targetRect.TopLeft, rectangles) Then
            Return True
        End If

        If IsPointVisible(fromPoint, targetRect.TopRight, rectangles) Then
            Return True
        End If

        If IsPointVisible(fromPoint, targetRect.BottomLeft, rectangles) Then
            Return True
        End If

        If IsPointVisible(fromPoint, targetRect.BottomRight, rectangles) Then
            Return True
        End If

        Return False
    End Function

    Private Shared Function RectangleIntersectsLine(ByVal rect As Rect, ByVal startPoint As Point, ByVal endPoint As Point) As Boolean
        rect.Inflate(-1, -1)
        Return rect.IntersectsWith(New Rect(startPoint, endPoint))
    End Function

    Private Shared Sub GetOppositeCorners(ByVal orientation As ConnectorOrientation, ByVal rect As Rect, ByRef n1 As Point, ByRef n2 As Point)
        Select Case orientation
            Case ConnectorOrientation.Left
                n1 = rect.TopRight
                n2 = rect.BottomRight
            Case ConnectorOrientation.Top
                n1 = rect.BottomLeft
                n2 = rect.BottomRight
            Case ConnectorOrientation.Right
                n1 = rect.TopLeft
                n2 = rect.BottomLeft
            Case ConnectorOrientation.Bottom
                n1 = rect.TopLeft
                n2 = rect.TopRight
            Case Else
                Throw New Exception("No opposite corners found!")
        End Select
    End Sub

    Private Shared Sub GetNeighborCorners(ByVal orientation As ConnectorOrientation, ByVal rect As Rect, ByRef n1 As Point, ByRef n2 As Point)
        Select Case orientation
            Case ConnectorOrientation.Left
                n1 = rect.TopLeft
                n2 = rect.BottomLeft
            Case ConnectorOrientation.Top
                n1 = rect.TopLeft
                n2 = rect.TopRight
            Case ConnectorOrientation.Right
                n1 = rect.TopRight
                n2 = rect.BottomRight
            Case ConnectorOrientation.Bottom
                n1 = rect.BottomLeft
                n2 = rect.BottomRight
            Case Else
                Throw New Exception("No neighour corners found!")
        End Select
    End Sub

    Private Shared Function Distance(ByVal p1 As Point, ByVal p2 As Point) As Double
        Return Point.Subtract(p1, p2).Length
    End Function

    Private Shared Function GetRectWithMargin(ByVal connectorThumb As ConnectorInfo, ByVal margin As Double) As Rect
        Dim rect = connectorThumb.DiagramRect
        rect.Inflate(margin, margin)
        Return rect
    End Function

    Private Shared Function GetOffsetPoint(ByVal connector As ConnectorInfo, ByVal rect As Rect) As Point
        Dim offsetPoint As New Point()

        Select Case connector.Orientation
            Case ConnectorOrientation.Left
                offsetPoint = New Point(rect.Left, connector.Position.Y)
            Case ConnectorOrientation.Top
                offsetPoint = New Point(connector.Position.X, rect.Top)
            Case ConnectorOrientation.Right
                offsetPoint = New Point(rect.Right, connector.Position.Y)
            Case ConnectorOrientation.Bottom
                offsetPoint = New Point(connector.Position.X, rect.Bottom)
            Case Else
        End Select

        Return offsetPoint
    End Function

    Private Shared Sub CheckPathEnd(ByVal source As ConnectorInfo, ByVal Target As ConnectorInfo, ByVal showLastLine As Boolean, ByVal linePoints As List(Of Point))
        If showLastLine Then
            Dim startPoint As New Point(0, 0)
            Dim endPoint As New Point(0, 0)
            Dim marginPath As Double = 10

            Select Case source.Orientation
                Case ConnectorOrientation.Left
                    startPoint = New Point(source.Position.X - marginPath, source.Position.Y)
                Case ConnectorOrientation.Top
                    startPoint = New Point(source.Position.X, source.Position.Y - marginPath)
                Case ConnectorOrientation.Right
                    startPoint = New Point(source.Position.X + marginPath, source.Position.Y)
                Case ConnectorOrientation.Bottom
                    startPoint = New Point(source.Position.X, source.Position.Y + marginPath)
                Case Else
            End Select

            marginPath = 20
            Select Case Target.Orientation
                Case ConnectorOrientation.Left
                    endPoint = New Point(Target.Position.X - marginPath, Target.Position.Y)
                Case ConnectorOrientation.Top
                    endPoint = New Point(Target.Position.X, Target.Position.Y - marginPath)
                Case ConnectorOrientation.Right
                    endPoint = New Point(Target.Position.X + marginPath, Target.Position.Y)
                Case ConnectorOrientation.Bottom
                    endPoint = New Point(Target.Position.X, Target.Position.Y + marginPath)
                Case Else
            End Select
            linePoints.Insert(0, startPoint)
            linePoints.Add(endPoint)
        Else
            linePoints.Insert(0, source.Position)
            linePoints.Add(Target.Position)
        End If
    End Sub

    Private Shared Function GetOpositeOrientation(ByVal connectorOrientation As ConnectorOrientation) As ConnectorOrientation
        Select Case connectorOrientation
            Case connectorOrientation.Left
                Return connectorOrientation.Right
            Case connectorOrientation.Top
                Return connectorOrientation.Bottom
            Case connectorOrientation.Right
                Return connectorOrientation.Left
            Case connectorOrientation.Bottom
                Return connectorOrientation.Top
            Case Else
                Return connectorOrientation.Top
        End Select
    End Function
End Class
