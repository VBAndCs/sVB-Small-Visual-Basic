Imports System.Threading
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Media
Imports System.Windows.Media.Animation
Imports System.Windows.Media.Imaging
Imports System.Windows.Shapes
Imports System.Windows.Threading
Imports Microsoft.SmallVisualBasic.Library.Internal

Namespace Library
    ''' <summary>
    ''' The Turtle provides Logo-like functionality to draw shapes by manipulating the properties of a pen and drawing primitives.
    ''' </summary>
    <SmallVisualBasicType>
    Public NotInheritable Class Turtle
        Private Const _turtleName As String = "_turtle"
        Friend Shared _initialized As Boolean
        Private Shared _isVisible As Boolean
        Private Shared _currentX As Double
        Private Shared _currentY As Double
        Private Shared _turtle As FrameworkElement
        Private Shared _speed As Integer
        Private Shared _angle As Double
        Private Shared _rotateTransform As RotateTransform
        Private Shared _penDown As Boolean
        Private Shared _toShow As Boolean
        Private Shared _width As Double = 16
        Private Shared _height As Double = 16

        ''' <summary>
        ''' Gets or sets the turtle width. The default value is 16, and the minimum value is 4.
        ''' Note that changeing the turtle width will also set its height to the same value.
        ''' </summary>     
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Property Width As Primitive
            Get
                Return _Width
            End Get

            Set(value As Primitive)
                ChangeSize(value)
            End Set
        End Property

        ''' <summary>
        ''' Gets or sets the turtle height. The default value is 16, and the minimum value is 4.
        ''' Note that changeing the turtle height will also set its width to the same value.
        ''' </summary>     
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Property Height As Primitive
            Get
                Return _height
            End Get

            Set(value As Primitive)
                ChangeSize(value)
            End Set
        End Property

        Private Shared Sub ChangeSize(value As Primitive)
            If Not value.IsNumber Then Return
            Dim v = System.Math.Abs(value.AsDecimal)
            If v < 4 Then v = 4
            _width = v
            _height = v
            VerifyAccess()

            SmallBasicApplication.Invoke(
                Sub()
                    _turtle.Width = v
                    _turtle.Height = v
                    v /= 2
                    _turtle.Margin = New Thickness(-v, -v, 0.0, 0.0)
                    _rotateTransform.CenterX = v
                    _rotateTransform.CenterY = v
                End Sub)
        End Sub

        Friend Shared Sub Initialize()
            _initialized = False
            _isVisible = True
            _currentX = 320.0
            _currentY = 240.0
            _speed = 15
            _angle = 0.0
            _width = 16.0
            _height = 16.0
            _penDown = True
            _toShow = False
            _turtle = Nothing
        End Sub

        ''' <summary>
        ''' Specifies how fast the turtle should move. 
        ''' Valid values are 1 to 50.
        ''' The default value is 15.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Property Speed As Primitive
            Get
                Return _speed
            End Get

            Set(Value As Primitive)
                _speed = Value
                If _speed < 1 Then
                    _speed = 1
                ElseIf _speed > 50 Then
                    _speed = 50
                End If
            End Set
        End Property

        ''' <summary>
        ''' Gets or sets the current angle of the turtle.  While setting, this will turn the turtle instantly to the new angle.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Property Angle As Primitive
            Get
                Return _angle
            End Get

            Set(Value As Primitive)
                _angle = Value
                If _rotateTransform IsNot Nothing Then
                    GraphicsWindow.Invoke(
                        Sub()
                            Dim animation As New DoubleAnimation With {
                                 .To = _angle,
                                 .Duration = New Duration(TimeSpan.FromMilliseconds(0.0))
                            }
                            _rotateTransform.BeginAnimation(RotateTransform.AngleProperty, animation)
                        End Sub)
                End If
            End Set
        End Property

        ''' <summary>
        ''' Gets or sets the X location of the Turtle.  While setting, this will move the turtle instantly to the new location.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Property X As Primitive
            Get
                Return _currentX
            End Get

            Set(Value As Primitive)
                _currentX = Value
                Shapes.Move(_turtleName, _currentX, _currentY)
            End Set
        End Property

        ''' <summary>
        ''' Gets or sets the Y location of the Turtle.  While setting, this will move the turtle instantly to the new location.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Property Y As Primitive
            Get
                Return _currentY
            End Get

            Set(Value As Primitive)
                _currentY = Value
                Shapes.Move(_turtleName, _currentX, _currentY)
            End Set
        End Property

        ''' <summary>
        ''' Shows the Turtle to enable interactions with it.
        ''' </summary>
        Public Shared Sub Show()
            _isVisible = True
            _toShow = True
            VerifyAccess()
            _toShow = False
        End Sub

        ''' <summary>
        ''' Hides the Turtle and disables interactions with it.
        ''' </summary>
        Public Shared Sub Hide()
            If _isVisible Then
                GraphicsWindow.VerifyAccess()
                GraphicsWindow.Invoke(Sub() Shapes.Remove(_turtleName))
                _isVisible = False
            End If
        End Sub

        ''' <summary>
        ''' Sets the pen down to enable the turtle to draw as it moves.
        ''' </summary>
        Public Shared Sub PenDown()
            _penDown = True
        End Sub

        ''' <summary>
        ''' Lifts the pen up to stop drawing as the turtle moves.
        ''' </summary>
        Public Shared Sub PenUp()
            _penDown = False
        End Sub

        Friend Shared _path As Path = Nothing
        Private Shared _group As GeometryGroup
        Private Shared _geometry As PathGeometry
        Private Shared _figure As PathFigure

        Shared Sub New()
            Initialize()
        End Sub

        ''' <summary>
        ''' Uses the next turtle movements to create a closed figure, so that you can fill it by calling the FillFigure method.
        ''' </summary>
        Public Shared Sub CreateFigure()
            GraphicsWindow.Invoke(
                Sub()
                    _group = New GeometryGroup()
                    _geometry = New PathGeometry()
                    _geometry.Figures = New PathFigureCollection()
                    _group.Children.Add(_geometry)
                    _path = New Path
                    _path.Data = _group
                    _figure = New PathFigure With {
                        .IsClosed = False,
                        .StartPoint = New Point(_currentX, _currentY)
                    }
                    _geometry.Figures.Add(_figure)
                End Sub)
        End Sub


        ''' <summary>
        ''' Closes the figure the Turtle created after calling the CreateFigure method, and filsl it with the GraphicsWindow.BrushColor.
        ''' After calling this method, the figure Is completed, And you need To create a New figure If you want To fill a New area.
        ''' If there Is no figure, calling this method will Do Nothing.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Sub FillFigure()
            If _path Is Nothing Then Return

            Dim name = Shapes.GenerateNewName("GeoPath")
            GraphicsWindow.Invoke(
                Sub()
                    GraphicsWindow.VerifyAccess()
                    _path.Fill = WinForms.Color.GetBrush(GraphicsWindow.BrushColor)
                    _path.Stroke = WinForms.Color.GetBrush(GraphicsWindow.PenColor)
                    _path.StrokeThickness = GraphicsWindow.PenWidth
                    GraphicsWindow.AddShape(name, _path)
                    _path = Nothing
                End Sub)
        End Sub

        ''' <summary>
        ''' Moves the turtle to a specified distance. If the pen is down, it will draw a line as it moves.
        ''' </summary>
        ''' <param name="distance">
        ''' The distance to move the turtle.
        ''' </param>
        Public Shared Sub Move(distance As Primitive)
            VerifyAccess()
            Dim d = CDbl(distance)
            Dim animateTime = 1
            If _speed < 50 Then
                animateTime = Math.Max(1, System.Math.Abs(d * 320.0 / (_speed ^ 2)))
            End If

            Dim angle = _angle / 180.0 * System.Math.PI
            Dim newY = _currentY - d * System.Math.Cos(angle)
            Dim newX = _currentX + d * System.Math.Sin(angle)
            Shapes.Animate(_turtleName, newX, newY, animateTime)

            GraphicsWindow.Invoke(
                Sub()
                    If _path IsNot Nothing Then
                        _figure.Segments.Add(New LineSegment(
                                New Point(newX, newY), _penDown
                            )
                        )
                    End If
                End Sub)

            If _penDown Then
                        GraphicsWindow.Invoke(
                    Sub()
                        Dim name As String = Shapes.GenerateNewName("_turtleLine")
                        Dim line1 As New Line With {
                              .Name = name,
                              .X1 = _currentX,
                              .Y1 = _currentY,
                              .Stroke = GraphicsWindow._pen.Brush,
                              .StrokeThickness = GraphicsWindow._pen.Thickness
                        }
                        GraphicsWindow.AddShape(name, line1)
                        Dim animation As New DoubleAnimation With {
                               .From = _currentX,
                               .To = newX,
                               .Duration = New Duration(TimeSpan.FromMilliseconds(animateTime))
                        }
                        Dim animation2 As New DoubleAnimation With {
                               .From = _currentY,
                               .To = newY,
                               .Duration = New Duration(TimeSpan.FromMilliseconds(animateTime))
                        }
                        line1.BeginAnimation(Line.X2Property, animation)
                        line1.BeginAnimation(Line.Y2Property, animation2)
                    End Sub)
                    End If

                    _currentX = newX
                    _currentY = newY
                    WaitForReturn(animateTime)
                End Sub

        ''' <summary>
        ''' Turns and moves the turtle to the specified location.  If the pen is down, it will draw a line as it moves.
        ''' </summary>
        ''' <param name="newX">
        ''' The x co-ordinate of the destination point.
        ''' </param>
        ''' <param name="newY">
        ''' The y co-ordinate of the destination point.
        ''' </param>
        Public Shared Sub MoveTo(newX As Primitive, newY As Primitive)
            Dim x1 = CDbl(_currentX)
            Dim x2 = CDbl(newX)
            Dim y1 = CDbl(_currentY)
            Dim y2 = CDbl(newY)

            Dim d = (x2 - x1) ^ 2 + (y2 - y1) ^ 2
            If d <> 0.0 Then
                Dim distance = System.Math.Sqrt(d)
                Dim delta = y1 - y2
                Dim angle = System.Math.Acos(delta / distance) * 180.0 / System.Math.PI
                If x2 < x1 Then
                    angle = 360.0 - angle
                End If

                Dim deltaAngle = angle - CDbl(Turtle.Angle) Mod 360
                If deltaAngle > 180.0 Then
                    deltaAngle -= 360.0
                End If

                Turn(deltaAngle)
                Move(distance)
            End If
        End Sub

        ''' <summary>
        ''' Turns the turtle by the specified angle.  Angle is in degrees and can be either positive or negative.  If the angle is positive, the turtle turns to its right.  If it is negative, the turtle turns to its left.
        ''' </summary>
        ''' <param name="angle">
        ''' The angle to turn the turtle.
        ''' </param>
        Public Shared Sub Turn(angle As Primitive)
            VerifyAccess()
            Dim a = CDbl(angle)
            Dim animateTime = If(_speed = 10, 1.0, System.Math.Abs(a * 200.0 / (_speed ^ 2)))
            _angle += a

            GraphicsWindow.Invoke(
                Sub()
                    Dim animation As New DoubleAnimation With {
                           .To = _angle,
                           .Duration = New Duration(TimeSpan.FromMilliseconds(animateTime))
                    }
                    _rotateTransform.BeginAnimation(RotateTransform.AngleProperty, animation)
                End Sub)
            WaitForReturn(animateTime)
        End Sub

        ''' <summary>
        ''' Turns the turtle 90 degrees to the right.
        ''' </summary>
        Public Shared Sub TurnRight()
            Turn(90)
        End Sub

        ''' <summary>
        ''' Turns the turtle 90 degrees to the left.
        ''' </summary>
        Public Shared Sub TurnLeft()
            Turn(-90)
        End Sub

        Private Shared Sub VerifyAccess()
            GraphicsWindow.VerifyAccess()
            If _initialized AndAlso Not _toShow Then Return
            _initialized = True
            If Not _isVisible Then Return

            SmallBasicApplication.Invoke(
                Sub()
                    If _turtle Is Nothing Then
                        Dim source As ImageSource = New BitmapImage(
                            SmallBasicApplication.GetResourceUri("Turtle.png")
                        )

                        _turtle = New Image With {
                              .Source = source,
                              .Margin = New Thickness(-_width / 2, -_height / 2, 0.0, 0.0),
                              .Height = _height,
                              .Width = _width
                        }

                        Panel.SetZIndex(_turtle, 1000000)
                        _rotateTransform = New RotateTransform With {
                                 .Angle = _angle,
                                 .CenterX = _width / 2,
                                 .CenterY = _height / 2
                          }
                        _turtle.RenderTransform = _rotateTransform
                    End If

                    Canvas.SetLeft(_turtle, _currentX)
                    Canvas.SetTop(_turtle, _currentY)
                    GraphicsWindow.AddShape(_turtleName, _turtle)
                End Sub)
        End Sub

        Private Shared Sub WaitForReturn(time As Double)
            If SmallBasicApplication.HasShutdown Then
                Return
            End If

            Dim evt As New AutoResetEvent(initialState:=False)
            SmallBasicApplication.Invoke(
                Sub()
                    Dim dt As New DispatcherTimer
                    dt.Interval = TimeSpan.FromMilliseconds(time)
                    AddHandler dt.Tick,
                    Sub()
                        evt.Set()
                        dt.Stop()
                    End Sub
                    dt.Start()
                End Sub)

            Dim millisecondsTimeout As Integer = 100
            If SmallBasicApplication.Dispatcher.CheckAccess() Then
                millisecondsTimeout = 10
            End If
            While Not evt.WaitOne(millisecondsTimeout) AndAlso Not SmallBasicApplication.HasShutdown
                SmallBasicApplication.ClearDispatcherQueue()
            End While
        End Sub
    End Class
End Namespace
