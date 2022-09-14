Imports System.Threading
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Media
Imports System.Windows.Media.Animation
Imports System.Windows.Media.Imaging
Imports System.Windows.Shapes
Imports System.Windows.Threading
Imports Microsoft.SmallBasic.Library.Internal

Namespace Library
    ''' <summary>
    ''' The Turtle provides Logo-like functionality to draw shapes by manipulating the properties of a pen and drawing primitives.
    ''' </summary>
    <SmallBasicType>
    Public NotInheritable Class Turtle
        Private Const _turtleName As String = "_turtle"
        Private Shared _initialized As Boolean = False
        Private Shared _isVisible As Boolean = True
        Private Shared _currentX As Double = 320.0
        Private Shared _currentY As Double = 240.0
        Private Shared _turtle As FrameworkElement
        Private Shared _speed As Integer = 5
        Private Shared _angle As Double = 0.0
        Private Shared _rotateTransform As RotateTransform
        Private Shared _penDown As Boolean = True
        Private Shared _toShow As Boolean = False

        ''' <summary>
        ''' Specifies how fast the turtle should move.  Valid values are 1 to 10.  If Speed is set to 10, the turtle moves and rotates instantly.
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
                ElseIf _speed > 10 Then
                    _speed = 10
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
                Shapes.Move("_turtle", _currentX, _currentY)
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
                Shapes.Move("_turtle", _currentX, _currentY)
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
                GraphicsWindow.Invoke(Sub() Shapes.Remove("_turtle"))
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

        ''' <summary>
        ''' Moves the turtle to a specified distance.  If the pen is down, it will draw a line as it moves.
        ''' </summary>
        ''' <param name="distance">
        ''' The distance to move the turtle.
        ''' </param>
        Public Shared Sub Move(distance As Primitive)
            VerifyAccess()
            Dim d = CDbl(distance)
            Dim animateTime = If(_speed = 10, 1.0, System.Math.Abs(d * 320.0 / (_speed ^ 2)))
            Dim angle = _angle / 180.0 * System.Math.PI
            Dim newY = _currentY - d * System.Math.Cos(angle)
            Dim newX = _currentX + d * System.Math.Sin(angle)
            Shapes.Animate("_turtle", newX, newY, animateTime)
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

            GraphicsWindow.Invoke(
                Sub()
                    If _turtle Is Nothing Then
                        Dim source As ImageSource = New BitmapImage(SmallBasicApplication.GetResourceUri("Turtle.png"))
                        _turtle = New Image With {
                              .Source = source,
                              .Margin = New Thickness(-8.0, -8.0, 0.0, 0.0),
                              .Height = 16.0,
                              .Width = 16.0
                        }

                        Panel.SetZIndex(_turtle, 1000000)
                        _rotateTransform = New RotateTransform With {
                                 .Angle = _angle,
                                 .CenterX = 8.0,
                                 .CenterY = 8.0
                          }
                        _turtle.RenderTransform = _rotateTransform
                        _turtle.Width = 16.0
                        _turtle.Height = 16.0
                    End If
                    Canvas.SetLeft(_turtle, _currentX)
                    Canvas.SetTop(_turtle, _currentY)
                    GraphicsWindow.AddShape("_turtle", _turtle)
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
