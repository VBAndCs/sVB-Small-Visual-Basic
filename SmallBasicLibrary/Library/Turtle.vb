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
        Private Shared _initialized As Boolean
        Friend Shared IsVisible As Boolean
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
                Return _width
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

        Friend Shared Sub Initialize(Optional resetSpeed As Boolean = True)
            _initialized = False
            IsVisible = True
            _currentX = 320.0
            _currentY = 240.0
            _angle = 0.0
            _width = 16.0
            _height = 16.0
            _penDown = True
            _toShow = False
            _turtle = Nothing
            If resetSpeed Then
                _speed = 5
                _useAnimation = True
            End If
        End Sub

        ''' <summary>
        ''' Specifies how fast the turtle should move. 
        ''' Valid values are 1 to 50. The default value is 5.
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

        Private Shared _useAnimation As Boolean = True

        ''' <summary>
        ''' Indicates whether or not the turtle moves and turns are animated.
        ''' The default value is True, which is OK in most cases, and you can increase the Turtle.Speed up to 50 to make it faster.
        ''' But you may need to set this property fo false while you are drawing curves, which require many turns and succissive short movements, which make the turtle too slow because of the animations overhead. 
        ''' Setting this property to False, will make the Move, MoveTo and Turn methods call the DirectMove, DirectMoveTo and DirectTurn Methods to avoid using animation, which will make thee turtle super fast.
        ''' You can also call the three `Direct` methods manually even when this property is set to True, which allows you to mix animated and non-animated moves.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Boolean)>
        Public Shared Property UseAnimation As Primitive
            Get
                Return New Primitive(_useAnimation)
            End Get

            Set(value As Primitive)
                _useAnimation = CBool(value)
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
            IsVisible = True
            _toShow = True
            VerifyAccess()
            _toShow = False
        End Sub

        ''' <summary>
        ''' Hides the Turtle and disables interactions with it.
        ''' </summary>
        Public Shared Sub Hide()
            If IsVisible Then
                GraphicsWindow.VerifyAccess()
                GraphicsWindow.Invoke(Sub() Shapes.Remove(_turtleName))
                IsVisible = False
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

        Private Shared animationX, animationY As DoubleAnimation

        ''' <summary>
        ''' Moves the turtle to a specified distance. If the pen is down, it will draw a line as it moves.
        ''' </summary>
        ''' <param name="distance">
        ''' The distance to move the turtle.
        ''' </param>
        Public Shared Sub Move(distance As Primitive)
            If Not _UseAnimation Then
                DirectMove(distance)
                Return
            End If

            VerifyAccess()
            If Not GraphicsWindow._windowCreated Then Return

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
                        If _penDown Then
                            Dim name As String = Shapes.GenerateNewName("_turtleLine")
                            Dim line1 As New Line With {
                                  .Name = name,
                                  .X1 = _currentX,
                                  .Y1 = _currentY,
                                  .Stroke = GraphicsWindow._pen.Brush,
                                  .StrokeThickness = GraphicsWindow._pen.Thickness
                            }
                            GraphicsWindow.AddShape(name, line1)

                            If animationX Is Nothing Then
                                animationX = New DoubleAnimation()
                                animationY = New DoubleAnimation()
                            End If

                            animationX.From = _currentX
                            animationX.To = newX
                            animationX.Duration = New Duration(TimeSpan.FromMilliseconds(animateTime))

                            animationY.From = _currentY
                            animationY.To = newY
                            animationY.Duration = New Duration(TimeSpan.FromMilliseconds(animateTime))

                            line1.BeginAnimation(Line.X2Property, animationX)
                            line1.BeginAnimation(Line.Y2Property, animationY)
                        End If

                        If _path IsNot Nothing Then
                            _figure.Segments.Add(New LineSegment(
                                    New Point(newX, newY), _penDown
                                )
                            )
                        End If
                    End Sub)

            _currentX = newX
            _currentY = newY
            WaitForReturn(animateTime)
        End Sub

        ''' <summary>
        ''' Moves the turtle to a specified distance directly without animation. If the pen is down, it will draw a line as it moves.
        ''' </summary>
        ''' <param name="distance">
        ''' The distance to move the turtle.
        ''' </param>
        Public Shared Sub DirectMove(distance As Primitive)
            VerifyAccess()
            If Not GraphicsWindow._windowCreated Then Return

            Dim d = CDbl(distance)
            Dim angle = _angle / 180.0 * System.Math.PI
            Dim newY = _currentY - d * System.Math.Cos(angle)
            Dim newX = _currentX + d * System.Math.Sin(angle)

            GraphicsWindow.Invoke(
                Sub()
                    If _penDown Then
                        Dim name = Shapes.GenerateNewName("_turtleLine")
                        Dim line1 As New Line With {
                              .Name = name,
                              .X1 = _currentX,
                              .X2 = newX,
                              .Y1 = _currentY,
                              .Y2 = newY,
                              .Stroke = GraphicsWindow._pen.Brush,
                              .StrokeThickness = GraphicsWindow._pen.Thickness
                        }

                        Shapes.Move(_turtleName, newX, newY)
                        GraphicsWindow.AddShape(name, line1)
                    End If

                    If _path IsNot Nothing Then
                        _figure.Segments.Add(New LineSegment(
                                New Point(newX, newX), _penDown))
                    End If
                End Sub)

            _currentX = newX
            _currentY = newY
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
            If Not _UseAnimation Then
                DirectMoveTo(newX, newY)
                Return
            End If

            Dim deltaAngle As Double
            Dim distance As Double
            GetAngleAndDistance(newX, newY, deltaAngle, distance)
            Turn(deltaAngle)
            Move(distance)
        End Sub

        ''' <summary>
        ''' Turns and moves the turtle to the specified location directly without any animation.  If the pen is down, it will draw a line as it moves.
        ''' </summary>
        ''' <param name="newX">
        ''' The x co-ordinate of the destination point.
        ''' </param>
        ''' <param name="newY">
        ''' The y co-ordinate of the destination point.
        ''' </param>
        Public Shared Sub DirectMoveTo(newX As Primitive, newY As Primitive)
            Dim deltaAngle As Double
            Dim distance As Double
            GetAngleAndDistance(newX, newY, deltaAngle, distance)
            DirectTurn(deltaAngle)
            DirectMove(distance)
        End Sub

        Private Shared Sub GetAngleAndDistance(
                    newX As Primitive,
                    newY As Primitive,
                    ByRef deltaAngle As Double,
                    ByRef distance As Double)

            Dim x1 = CDbl(_currentX)
            Dim x2 = CDbl(newX)
            Dim y1 = CDbl(_currentY)
            Dim y2 = CDbl(newY)

            Dim d = (x2 - x1) ^ 2 + (y2 - y1) ^ 2
            If d <> 0.0 Then
                distance = System.Math.Sqrt(d)
                Dim delta = y1 - y2
                Dim angle = System.Math.Acos(delta / distance) * 180.0 / System.Math.PI
                If x2 < x1 Then
                    angle = 360.0 - angle
                End If

                deltaAngle = angle - CDbl(Turtle.Angle) Mod 360
                If deltaAngle > 180.0 Then
                    deltaAngle -= 360.0
                End If
            End If
        End Sub

        Private Shared animationA As DoubleAnimation

        ''' <summary>
        ''' Turns the turtle by the specified angle.
        ''' </summary>
        ''' <param name="angle">
        ''' The angle in degrees to turn the turtle. It can be either positive or negative:
        ''' If the angle is positive, the turtle turns to its right.
        ''' If it is negative, the turtle turns to its left.
        ''' </param>
        Public Shared Sub Turn(angle As Primitive)
            If Not _UseAnimation Then
                DirectTurn(angle)
                Return
            End If

            VerifyAccess()
            If Not GraphicsWindow._windowCreated Then Return

            Dim a = CDbl(angle)
            Dim animateTime = If(_speed = 10, 1.0, System.Math.Abs(a * 200.0 / (_speed ^ 2)))
            _angle += a

            GraphicsWindow.Invoke(
                Sub()
                    If animationA Is Nothing Then animationA = New DoubleAnimation
                    animationA.To = _angle
                    animationA.Duration = New Duration(TimeSpan.FromMilliseconds(animateTime))
                    _rotateTransform.BeginAnimation(RotateTransform.AngleProperty, animationA)
                End Sub)
            WaitForReturn(animateTime)
        End Sub

        ''' <summary>
        ''' Turns the turtle by the specified angle directly without any animation.        
        ''' </summary>
        ''' <param name="angle">
        ''' The angle in degrees to turn the turtle. It can be either positive or negative:
        ''' If the angle is positive, the turtle turns to its right.
        ''' If it is negative, the turtle turns to its left.
        ''' </param>
        Public Shared Sub DirectTurn(angle As Primitive)
            VerifyAccess()
            If Not GraphicsWindow._windowCreated Then Return

            _angle += CDbl(angle)

            GraphicsWindow.Invoke(
                Sub() _rotateTransform.Angle = _angle
            )
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
            If Not IsVisible Then Return

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
