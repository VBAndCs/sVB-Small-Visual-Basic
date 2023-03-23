Imports System.Windows
Imports System.Windows.Media
Imports System.Windows.Shapes
Imports Microsoft.SmallVisualBasic.Library

Namespace WinForms

    ''' <summary>
    ''' Allows you to combine shapes into one path, and create new geometric figures from line and arc segments, so you can create new complex custom shapes.
    ''' You can add the geometric path to the shapes collection by calling the Shapes.AddGeometricPath() method, then you can apply any rotation or animation on it as you do with any other normal shape.
    ''' You can also add the the geometric path to any label on any form by calling the Label.AddGeometricPath method, so you can move, rotate or animate it using the label methods.
    ''' </summary>
    <SmallVisualBasicType>
    Public Class GeometricPath
        Friend Shared _path As Path

        Private Shared _group As GeometryGroup
        Private Shared _geometry As PathGeometry
        Private Shared _figure As PathFigure

        Private Shared ReadOnly Property Group As GeometryGroup
            Get
                If _group Is Nothing Then CreatePath()
                Return _group
            End Get
        End Property

        Private Shared ReadOnly Property PathGeometry As PathGeometry
            Get
                If _geometry Is Nothing Then CreatePath()
                Return _geometry
            End Get
        End Property

        Private Shared ReadOnly Property PathFigure As PathFigure
            Get
                If _figure Is Nothing Then CreateFigure(0, 0, True)
                Return _figure
            End Get
        End Property

        ''' <summary>
        ''' creates a new geometric path to add geometric shapes to it, so you can compose complex shapes.
        ''' </summary>
        Public Shared Sub CreatePath()
            GraphicsWindow.Invoke(
                Sub()
                    _group = New GeometryGroup()
                    _geometry = New PathGeometry()
                    _geometry.Figures = New PathFigureCollection()
                    _group.Children.Add(_geometry)
                    _path = New Path
                    _path.Data = _group
                End Sub)
        End Sub

        ''' <summary>
        ''' Adds a rectangle to the geometric path.
        ''' </summary>
        ''' <param name="x">The x co-ordinate of the rectangle.</param>
        ''' <param name="y">The y co-ordinate of the rectangle.</param>
        ''' <param name="width">The width of the rectangle.</param>
        ''' <param name="height">The height of the rectangle.</param>
        Public Shared Sub AddRectangle(x As Primitive, y As Primitive, width As Primitive, height As Primitive)
            GraphicsWindow.Invoke(
                Sub()
                    Dim r As New Rect(CInt(x), CInt(y), width, height)
                    Group.Children.Add(New RectangleGeometry(r))
                End Sub)
        End Sub


        ''' <summary>
        ''' Adds an ellipse to the geometric path
        ''' </summary>
        ''' <param name="x">The x co-ordinate of the ellipse.</param>
        ''' <param name="y">The y co-ordinate of the ellipse.</param>
        ''' <param name="width">The width of the ellipse.</param>
        ''' <param name="height">The height of the ellipse.</param>
        Public Shared Sub AddEllipse(x As Primitive, y As Primitive, width As Primitive, height As Primitive)
            GraphicsWindow.Invoke(
                Sub()
                    Dim w = CDbl(width) / 2
                    Dim h = CDbl(height) / 2
                    Dim p As New Point(CDbl(x) + w, CDbl(y) + h)
                    Group.Children.Add(New EllipseGeometry(p, w, h))
                End Sub)
        End Sub

        ''' <summary>
        ''' Adds a triangle to the geometric path.
        ''' </summary>
        ''' <param name="x1">The x co-ordinate of the first point.</param>
        ''' <param name="y1">The y co-ordinate of the first point.</param>
        ''' <param name="x2">The x co-ordinate of the second point.</param>
        ''' <param name="y2">The y co-ordinate of the second point.</param>
        ''' <param name="x3">The x co-ordinate of the third point.</param>
        ''' <param name="y3">The y co-ordinate of the third point.</param>
        Public Shared Sub AddTriangle(
                          x1 As Primitive, y1 As Primitive,
                          x2 As Primitive, y2 As Primitive,
                          x3 As Primitive, y3 As Primitive
                   )

            GraphicsWindow.Invoke(
                Sub()
                    Dim figure As New PathFigure With {
                        .StartPoint = New Point(x1, y1),
                        .IsClosed = True,
                        .Segments = New PathSegmentCollection From {
                                New LineSegment(New Point(x2, y2), isStroked:=True),
                                New LineSegment(New Point(x3, y3), isStroked:=True)
                        }
                    }

                    Dim triangle As New PathGeometry
                    triangle.Figures = New PathFigureCollection From {figure}
                    triangle.Freeze()
                    Group.Children.Add(triangle)
                End Sub)
        End Sub

        ''' <summary>
        ''' Adds a line to the geometric path.
        ''' </summary>
        ''' <param name="x1">The x co-ordinate of the first point.</param>
        ''' <param name="y1">The y co-ordinate of the first point.</param>
        ''' <param name="x2">The x co-ordinate of the second point.</param>
        ''' <param name="y2">The y co-ordinate of the second point.</param>
        Public Shared Sub AddLine(x1 As Primitive, y1 As Primitive, x2 As Primitive, y2 As Primitive)
            GraphicsWindow.Invoke(
                Sub()
                    Group.Children.Add(New LineGeometry(
                            New Point(x1, y1),
                            New Point(x2, y2)
                        )
                    )
                End Sub)
        End Sub

        ''' <summary>
        ''' Adds a new figure to the geometric path that starts at the given point
        ''' </summary>
        ''' <param name="x">The x co-ordinate of the start point of the figure</param>
        ''' <param name="y">The y co-ordinate of the start point of the figure</param>
        ''' <param name="isClosed">When True, a line segment is automatically drown between the last point and the start point of the figure to make it a closed shape. Use Fals if you want to draw an open figure like a curve.</param>
        Public Shared Sub CreateFigure(x As Primitive, y As Primitive, isClosed As Primitive)
            GraphicsWindow.Invoke(
                Sub()
                    _figure = New PathFigure With {
                        .IsClosed = isClosed,
                        .StartPoint = New Point(x, y)
                    }
                    PathGeometry.Figures.Add(_figure)
                End Sub)
        End Sub

        ''' <summary>
        ''' Adds a line segment to the current figure in the geometric path, starting from the last point in the figure to the given point.
        ''' </summary>
        ''' <param name="x">The x co-ordinate of the end point of the line segment.</param>
        ''' <param name="y">The y co-ordinate of the end point of the line segment.</param>
        ''' <param name="usePen">Use True to draw the segment with the pen color, or False to hide the segment outline.</param>
        Public Shared Sub AddLineSegment(x As Primitive, y As Primitive, usePen As Primitive)
            GraphicsWindow.Invoke(
                Sub()
                    PathFigure.Segments.Add(New LineSegment(
                            New Point(x, y),
                            usePen
                        )
                    )
                End Sub)
        End Sub

        ''' <summary>
        ''' Adds an arc segment to the current figure in the geometric path, starting from the last point in the figure to the given point.
        ''' </summary>
        ''' <param name="x">The x co-ordinate of the end point of the arc segment.</param>
        ''' <param name="y">The y co-ordinate of the end point of the arc segment.</param>
        ''' <param name="xRadius">The horizontal radius of the arc</param>
        ''' <param name="yRadius">The vertical radius of the arc</param>
        ''' <param name="angle">The x-axis rotation of the ellipse</param>
        ''' <param name="isLargArc">Use True if the arc should be greater than 180 degrees, or False otherwise</param>
        ''' <param name="isClockwise">Use True to draw the arc in a positive angle direction, or False otherwise</param>
        ''' <param name="usePen">Use True to draw the segment with the pen color, or False to hide the segment outline.</param>
        Public Shared Sub AddArcSegment(
                       x As Primitive,
                       y As Primitive,
                       xRadius As Primitive,
                       yRadius As Primitive,
                       angle As Primitive,
                       isLargArc As Primitive,
                       isClockwise As Primitive,
                       usePen As Primitive
                   )

            GraphicsWindow.Invoke(
                Sub()
                    PathFigure.Segments.Add(New ArcSegment(
                            New Point(x, y),
                            New Size(xRadius, yRadius),
                            angle,
                            isLargArc,
                            If(isClockwise, SweepDirection.Clockwise, SweepDirection.Counterclockwise),
                            usePen
                        )
                    )
                End Sub)
        End Sub

        ''' <summary>
        ''' Adds a cubic Bezier curve segment to the current figure in the geometric path, starting from the last point in the figure, passing throw the two given, and ending at the given end point control points.
        ''' </summary>
        ''' <param name="x1">The x co-ordinate of the first control point.</param>
        ''' <param name="y1">The y co-ordinate of the first control point.</param>
        ''' <param name="x2">The x co-ordinate of the second control point.</param>
        ''' <param name="y2">The y co-ordinate of the second control point.</param>
        ''' <param name="x3">The x co-ordinate of the end point.</param>
        ''' <param name="y3">The y co-ordinate of the end point.</param>
        ''' <param name="usePen">Use True to draw the segment with the pen color, or False to hide the segment outline.</param>
        Public Shared Sub AddBezierSegment(
                         x1 As Primitive, y1 As Primitive,
                         x2 As Primitive, y2 As Primitive,
                         x3 As Primitive, y3 As Primitive,
                         usePen As Primitive
                   )

            GraphicsWindow.Invoke(
                Sub()
                    PathFigure.Segments.Add(New BezierSegment(
                            New Point(x1, y1),
                            New Point(x2, y2),
                            New Point(x3, y3),
                            usePen
                        )
                    )
                End Sub)
        End Sub

        ''' <summary>
        ''' Adds a cubic quadratic Bezier curve segment to the current figure in the geometric path, starting from the last point in the figure, passing throw the given control point and ending at the given end point.
        ''' </summary>
        ''' <param name="x1">The x co-ordinate of the control point.</param>
        ''' <param name="y1">The y co-ordinate of the control point.</param>
        ''' <param name="x2">The x co-ordinate of the end point.</param>
        ''' <param name="y2">The y co-ordinate of the end point.</param>
        ''' <param name="usePen">Use True to draw the segment with the pen color, or False to hide the segment outline.</param>
        Public Shared Sub AddQuadraticBezierSegment(
                         x1 As Primitive, y1 As Primitive,
                         x2 As Primitive, y2 As Primitive,
                         usePen As Primitive
                   )

            GraphicsWindow.Invoke(
                Sub()
                    PathFigure.Segments.Add(New QuadraticBezierSegment(
                            New Point(x1, y1),
                            New Point(x2, y2),
                            usePen
                        )
                    )
                End Sub)
        End Sub

    End Class
End Namespace
