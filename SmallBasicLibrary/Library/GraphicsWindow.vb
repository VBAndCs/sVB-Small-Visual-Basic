Imports System.ComponentModel
Imports System.Drawing
Imports System.Globalization
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input
Imports System.Windows.Interop
Imports System.Windows.Media
Imports System.Windows.Media.Animation
Imports System.Windows.Media.Imaging
Imports Microsoft.SmallVisualBasic.Library.Internal

Namespace Library
    ''' <summary>
    ''' The GraphicsWindow provides graphics related input and output functionality.  For example, using this class, it is possible to draw and fill circles and rectangles.
    ''' </summary>
    <SmallVisualBasicType>
    Public NotInheritable Class GraphicsWindow
        Private Class VisualContainer
            Inherits FrameworkElement
            Private _drawing As Drawing

            Public Sub New(drawing As Drawing)
                _drawing = drawing
            End Sub

            Protected Overrides Sub OnRender(dc As DrawingContext)
                dc.DrawDrawing(_drawing)
            End Sub
        End Class

        Private Const _defaultFontName As String = "Tahoma"
        Private Const _step As Integer = 20
        Friend Shared _windowVisible As Boolean = False
        Friend Shared _windowCreated As Boolean = False
        Friend Shared _window As Window
        Friend Shared _pen As Media.Pen
        Friend Shared _fillBrush As Media.Brush = Media.Brushes.SlateBlue
        Friend Shared _fontFamily As Media.FontFamily
        Friend Shared _fontSize As Double = 12.0
        Friend Shared _fontStyle As System.Windows.FontStyle = FontStyles.Normal
        Friend Shared _fontWeight As FontWeight = FontWeights.Bold
        Private Shared _mouseX As Double
        Private Shared _mouseY As Double
        Private Shared _lastKey As Key
        Private Shared _lastText As String
        Private Shared _random As Random
        Private Shared _backgroundColor As Primitive
        Friend Shared _objectsMap As New Dictionary(Of String, UIElement)
        Private Shared _scaleTransformMap As New Dictionary(Of String, ScaleTransform)
        Private Shared _mainCanvas As Canvas
        Private Shared _mainDrawing As DrawingGroup
        Private Shared _visualContainer As VisualContainer
        Private Shared _renderBitmap As RenderTargetBitmap
        Private Shared _bitmapContainer As System.Windows.Controls.Image
        Private Shared _isHidden As Boolean
        Private Shared _syncLock As New Object

        ''' <summary>
        ''' Gets or sets whether or not the Graphics Window is the top most window that always appears on top of all other desktop windows even when it is not the active window.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Boolean)>
        Public Shared Property Topmost As Primitive
            Get
                VerifyAccess()
                Invoke(Sub() Topmost = _window.Topmost)
            End Get

            Set(value As Primitive)
                VerifyAccess()
                Invoke(Sub() _window.Topmost = CBool(value))
            End Set
        End Property

        ''' <summary>
        ''' Gets or sets the Background color of the Graphics Window.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Color)>
        Public Shared Property BackgroundColor As Primitive
            Get
                VerifyAccess()
                Return _backgroundColor
            End Get

            Set(Value As Primitive)
                _backgroundColor = Value
                VerifyAccess(True)
                Invoke(
                    Sub()
                        _window.Background = WinForms.Color.GetBrush(Value)
                    End Sub)
            End Set
        End Property

        ''' <summary>
        ''' Fills the Graphics Window background with a gradient brush that starts with its background color and ends with the given color.
        ''' </summary>
        ''' <param name="endColor">The end color of the gradient brush</param>
        Public Shared Sub FillGradiantBackground(endColor As Primitive)
            VerifyAccess()
            Invoke(
                    Sub()
                        If WinForms.Color.IsNone(endColor) Then
                            ' Use a solid brush
                            _window.Background = WinForms.Color.GetBrush(_backgroundColor)
                        Else
                            _window.Background = New System.Windows.Media.LinearGradientBrush(
                                 WinForms.Color.FromString(_backgroundColor),
                                 WinForms.Color.FromString(endColor),
                                 New System.Windows.Point(0, 0),
                                 New System.Windows.Point(1, 0)
                            )
                        End If
                    End Sub)
        End Sub

        Private Shared _brushColor As Primitive = WinForms.Colors.CornflowerBlue
        Private Shared _gradientEndColor As Primitive = WinForms.Colors.None

        ''' <summary>
        ''' Gets or sets the brush color to be used to fill shapes drawn on the Graphics Window.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Color)>
        Public Shared Property BrushColor As Primitive
            Get
                VerifyAccess()
                Return _brushColor
            End Get

            Set(Value As Primitive)
                VerifyAccess()
                _brushColor = Value
                CreateBrush()
            End Set
        End Property

        ''' <summary>
        ''' Gets Or sets the end color for the gradient brush. Use the BrushColor property  to set the start color of this gradient brush.
        ''' The Default value Is Colors.None, which means that the BrushColor will be used alone To create a solid brush.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Color)>
        Public Shared Property GradientEndColor As Primitive
            Get
                VerifyAccess()
                Return _gradientEndColor
            End Get

            Set(Value As Primitive)
                VerifyAccess()
                _gradientEndColor = Value
                CreateBrush()
            End Set
        End Property

        Private Shared Sub CreateBrush()
            Invoke(
                Sub()
                    'Try
                    If WinForms.Color.IsNone(_gradientEndColor) Then
                            _fillBrush = WinForms.Color.GetBrush(_brushColor)
                        Else
                            _fillBrush = New System.Windows.Media.LinearGradientBrush(
                             WinForms.Color.FromString(_brushColor),
                             WinForms.Color.FromString(_gradientEndColor),
                             New System.Windows.Point(0, 0),
                             New System.Windows.Point(1, 0)
                        )
                        End If
                        _fillBrush?.Freeze()
                    'Catch
                    'End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Specifies whether or not the Graphics Window can be resized by the user.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Boolean)>
        Public Shared Property CanResize As Primitive
            Get
                VerifyAccess()
                Return InvokeWithReturn(
                     Function() As Primitive
                         Return New Primitive(_window.ResizeMode = ResizeMode.CanResize)
                     End Function)
            End Get

            Set(Value As Primitive)
                VerifyAccess()
                Invoke(
                    Sub()
                        If CBool(Value) Then
                            _window.ResizeMode = ResizeMode.CanResize
                        Else
                            _window.ResizeMode = ResizeMode.NoResize
                        End If
                    End Sub)
            End Set
        End Property

        Private Shared _penWidth As Primitive = New Primitive(2)

        ''' <summary>
        ''' Gets or sets the width of the pen used to draw shapes on the Graphics Window.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Property PenWidth As Primitive
            Get
                VerifyAccess()
                Return _penWidth
            End Get

            Set(Value As Primitive)
                VerifyAccess()
                _penWidth = Value
                CreatePen()
            End Set
        End Property

        Private Shared _penColor As Primitive = WinForms.Colors.Blue

        ''' <summary>
        ''' Gets or sets the color of the pen used to draw shapes on the Graphics Window.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Color)>
        Public Shared Property PenColor As Primitive
            Get
                VerifyAccess()
                Return _penColor
            End Get

            Set(Value As Primitive)
                VerifyAccess()
                _penColor = Value
                CreatePen()
            End Set
        End Property

        Private Shared Sub CreatePen()
            Invoke(
                Sub()
                    If WinForms.Color.IsNone(_penColor) Then
                        _pen = Nothing
                    Else
                        _pen = New Media.Pen(WinForms.Color.GetBrush(_penColor), _penWidth)
                        _pen.Freeze()
                    End If
                End Sub)
        End Sub

        ''' <summary>
        ''' Gets or sets the Font Name to be used when drawing text on the Graphics Window.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Property FontName As Primitive
            Get
                Return InvokeWithReturn(
                    Function() As Primitive
                        Return New Primitive(If(_fontFamily IsNot Nothing, _fontFamily.Source, "Tahoma"))
                    End Function)
            End Get

            Set(Value As Primitive)
                Invoke(Sub() _fontFamily = New Media.FontFamily(Value))
            End Set
        End Property

        ''' <summary>
        ''' Gets or sets the Font Size to be used when drawing text on the Graphics Window.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Property FontSize As Primitive
            Get
                Return New Primitive(_fontSize)
            End Get

            Set(Value As Primitive)
                Invoke(Sub() _fontSize = Value)
            End Set
        End Property

        ''' <summary>
        ''' Gets or sets whether or not the font to be used when drawing text on the Graphics Window, is bold.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Boolean)>
        Public Shared Property FontBold As Primitive
            Get
                Return _fontWeight = FontWeights.Bold
            End Get

            Set(Value As Primitive)
                Invoke(
                    Sub()
                        If CBool(Value) Then
                            _fontWeight = FontWeights.Bold
                        Else
                            _fontWeight = FontWeights.Normal
                        End If
                    End Sub)
            End Set
        End Property

        ''' <summary>
        ''' Gets or sets whether or not the font to be used when drawing text on the Graphics Window, is italic.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Boolean)>
        Public Shared Property FontItalic As Primitive
            Get
                Return _fontStyle = FontStyles.Italic
            End Get

            Set(Value As Primitive)
                Invoke(
                    Sub()
                        If CBool(Value) Then
                            _fontStyle = FontStyles.Italic
                        Else
                            _fontStyle = FontStyles.Normal
                        End If
                    End Sub)
            End Set
        End Property

        ''' <summary>
        ''' Gets or sets the title for the graphics window.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Property Title As Primitive
            Get
                VerifyAccess()
                Return InvokeWithReturn(
                    Function() As Primitive
                        Return New Primitive(_window.Title)
                    End Function)
            End Get

            Set(Value As Primitive)
                VerifyAccess()
                Invoke(Sub() _window.Title = Value)
            End Set
        End Property

        ''' <summary>
        ''' Gets or sets the Height of the graphics window.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Property Height As Primitive
            Get
                VerifyAccess()
                Return InvokeWithReturn(
                    Function() As Primitive
                        Return New Primitive(_mainCanvas.ActualHeight)
                    End Function)
            End Get

            Set(Value As Primitive)
                VerifyAccess()
                Invoke(
                    Sub()
                        _window.WindowState = WindowState.Normal
                        _window.Height = Value + (_window.ActualHeight - _mainCanvas.ActualHeight)
                    End Sub)
            End Set
        End Property

        ''' <summary>
        ''' Gets or sets the Width of the graphics window.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Property Width As Primitive
            Get
                VerifyAccess()
                Return InvokeWithReturn(
                    Function() As Primitive
                        Return New Primitive(_mainCanvas.ActualWidth)
                    End Function)
            End Get

            Set(Value As Primitive)
                VerifyAccess()
                Invoke(
                    Sub()
                        _window.WindowState = WindowState.Normal
                        _window.Width = Value + (_window.ActualWidth - _mainCanvas.ActualWidth)
                    End Sub)
            End Set
        End Property

        ''' <summary>
        ''' Returns True is the GW is closed, otherwise False.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Boolean)>
        Public Shared ReadOnly Property IsClosed As Primitive
            Get
                Return Not _windowCreated
            End Get
        End Property

        ''' <summary>
        ''' Set this property to False to prevent showing the Graphics Window when any of its methods is called.
        ''' The Default value is True
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Boolean)>
        Public Shared Property AutoShow As Primitive = True


        ''' <summary>
        ''' Gets or sets the Left Position of the graphics window.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Property Left As Primitive
            Get
                VerifyAccess()
                Return InvokeWithReturn(
                    Function() As Primitive
                        Return New Primitive(_window.Left)
                    End Function)
            End Get

            Set(Value As Primitive)
                VerifyAccess()
                Invoke(Sub() _window.Left = Value)
            End Set
        End Property

        ''' <summary>
        ''' Gets or sets the Top Position of the graphics window.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Property Top As Primitive
            Get
                VerifyAccess()
                Return InvokeWithReturn(
                    Function() As Primitive
                        Return New Primitive(_window.Top)
                    End Function)
            End Get

            Set(Value As Primitive)
                VerifyAccess()
                Invoke(Sub() _window.Top = Value)
            End Set
        End Property

        Private Shared windPosX As Double = 0
        Private Shared windPosY As Double = 0
        Private Shared windWidth As Double = 0
        Private Shared windHeight As Double = 0
        Private Shared _fullScreen As Primitive


        ''' <summary>
        ''' Set this property to True to show the Graphics Window in the full screen mode, or set it to False to exit the full screen mode and restore its normal state.
        ''' Note that the user can press F11 to toggle the value of this property.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Boolean)>
        Public Shared Property FullScreen As Primitive
            Get
                VerifyAccess()
                Return _fullScreen
            End Get

            Set(value As Primitive)
                VerifyAccess()
                _fullScreen = value
                Invoke(
                    Sub()
                        If CBool(value) Then
                            windPosX = _window.Left
                            windPosY = _window.Top
                            windWidth = _window.ActualWidth
                            windHeight = _window.ActualHeight
                            If Not _window.AllowsTransparency Then
                                _window.WindowStyle = WindowStyle.None
                                _window.ResizeMode = ResizeMode.NoResize
                            End If
                            _window.Left = 0
                            _window.Top = 0
                            _window.Width = SystemParameters.WorkArea.Width
                            _window.Height = SystemParameters.WorkArea.Height

                        ElseIf _window.WindowStyle = WindowStyle.None Then
                            If Not _window.AllowsTransparency Then
                                _window.WindowStyle = WindowStyle.ThreeDBorderWindow
                                _window.ResizeMode = ResizeMode.CanResize
                            End If
                            _window.Width = windWidth
                            _window.Height = windHeight
                            _window.Left = windPosX
                            _window.Top = windPosY
                        End If
                    End Sub)
            End Set
        End Property


        ''' <summary>
        ''' Gets the last key name that was pressed or released. 
        ''' You cant compare this property to the values of th Keys enum, which is valid only for the Event.LastKey and Keyboard.LastKey properties.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared ReadOnly Property LastKey As Primitive
            Get
                Return New Primitive(_lastKey.ToString())
            End Get
        End Property

        ''' <summary>
        ''' Gets the last text that was entered on the Graphics Window.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared ReadOnly Property LastText As Primitive
            Get
                Return New Primitive(_lastText)
            End Get
        End Property

        ''' <summary>
        ''' Gets the x-position of the mouse relative to the Graphics Window.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared ReadOnly Property MouseX As Primitive
            Get
                Return _mouseX
            End Get
        End Property

        ''' <summary>
        ''' Gets the y-position of the mouse relative to the Graphics Window.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared ReadOnly Property MouseY As Primitive
            Get
                Return _mouseY
            End Get
        End Property

        Private Shared keyDownHandler As SmallVisualBasicCallback

        ''' <summary>
        ''' Raises an event when a key is pressed down on the keyboard.
        ''' </summary>
        Public Shared Custom Event KeyDown As SmallVisualBasicCallback
            AddHandler(Value As SmallVisualBasicCallback)
                VerifyAccess()
                keyDownHandler = Value
            End AddHandler

            RemoveHandler(Value As SmallVisualBasicCallback)
                keyDownHandler = Nothing
            End RemoveHandler

            RaiseEvent()
                If keyDownHandler IsNot Nothing Then keyDownHandler.Invoke()
            End RaiseEvent
        End Event

        Private Shared keyUpHandler As SmallVisualBasicCallback

        ''' <summary>
        ''' Raises an event when a key is released on the keyboard.
        ''' </summary>
        Public Shared Custom Event KeyUp As SmallVisualBasicCallback
            AddHandler(Value As SmallVisualBasicCallback)
                VerifyAccess()
                keyUpHandler = Value
            End AddHandler

            RemoveHandler(Value As SmallVisualBasicCallback)
                keyUpHandler = Nothing
            End RemoveHandler

            RaiseEvent()
                If keyUpHandler IsNot Nothing Then keyUpHandler.Invoke()
            End RaiseEvent
        End Event

        Private Shared mouseDownHandler As SmallVisualBasicCallback

        ''' <summary>
        ''' Raises an event when the mouse button is clicked down.
        ''' </summary>
        Public Shared Custom Event MouseDown As SmallVisualBasicCallback
            AddHandler(Value As SmallVisualBasicCallback)
                VerifyAccess()
                mouseDownHandler = Value
            End AddHandler

            RemoveHandler(Value As SmallVisualBasicCallback)
                mouseDownHandler = Nothing
            End RemoveHandler

            RaiseEvent()
                If mouseDownHandler IsNot Nothing Then mouseDownHandler.Invoke()
            End RaiseEvent

        End Event

        Private Shared mouseUpHandler As SmallVisualBasicCallback

        ''' <summary>
        ''' Raises an event when the mouse button is released.
        ''' </summary>
        Public Shared Custom Event MouseUp As SmallVisualBasicCallback
            AddHandler(Value As SmallVisualBasicCallback)
                VerifyAccess()
                mouseUpHandler = Value
            End AddHandler

            RemoveHandler(Value As SmallVisualBasicCallback)
                mouseUpHandler = Nothing
            End RemoveHandler

            RaiseEvent()
                If mouseUpHandler IsNot Nothing Then mouseUpHandler.Invoke()
            End RaiseEvent
        End Event

        Private Shared mouseMoveHandler As SmallVisualBasicCallback

        ''' <summary>
        ''' Raises an event when the mouse is moved around.
        ''' </summary>
        Public Shared Custom Event MouseMove As SmallVisualBasicCallback
            AddHandler(Value As SmallVisualBasicCallback)
                VerifyAccess()
                mouseMoveHandler = Value
            End AddHandler

            RemoveHandler(Value As SmallVisualBasicCallback)
                mouseMoveHandler = Nothing
            End RemoveHandler

            RaiseEvent()
                If mouseMoveHandler IsNot Nothing Then mouseMoveHandler.Invoke()
            End RaiseEvent
        End Event

        Private Shared textInputHandler As SmallVisualBasicCallback

        ''' <summary>
        ''' Raises an event when text is entered on the GraphicsWindow.
        ''' </summary>
        Public Shared Custom Event TextInput As SmallVisualBasicCallback
            AddHandler(Value As SmallVisualBasicCallback)
                VerifyAccess()
                textInputHandler = Value
            End AddHandler

            RemoveHandler(Value As SmallVisualBasicCallback)
                textInputHandler = Nothing
            End RemoveHandler

            RaiseEvent()
                If textInputHandler IsNot Nothing Then textInputHandler.Invoke()
            End RaiseEvent
        End Event

        ''' <summary>
        ''' Shows the Graphics window to enable interactions with it.
        ''' </summary>
        Public Shared Sub Show()
            If Not _windowCreated Then
                CreateWindow()
                Invoke(Sub() _window.WindowState = WindowState.Maximized)
            End If

            _isHidden = False
            Invoke(
                Sub()
                    If _windowVisible Then
                        If _window.WindowState = WindowState.Minimized Then
                            _window.WindowState = WindowState.Normal
                        End If
                        _window.Activate()
                    Else
                        _window.Show()
                        _windowVisible = True
                    End If
                End Sub)
        End Sub

        ''' <summary>
        ''' Hides the Graphics window.
        ''' </summary>
        Public Shared Sub Hide()
            _isHidden = True
            If _windowVisible Then
                Invoke(Sub() _window.Hide())
                _windowVisible = False
            End If
        End Sub

        ''' <summary>
        ''' Draws a rectangle on the screen using the selected Pen.
        ''' </summary>
        ''' <param name="x">The x co-ordinate of the rectangle.</param>
        ''' <param name="y">The y co-ordinate of the rectangle.</param>
        ''' <param name="width">The width of the rectangle.</param>
        ''' <param name="height">The height of the rectangle.</param>
        Public Shared Sub DrawRectangle(x As Primitive, y As Primitive, width As Primitive, height As Primitive)
            VerifyAccess()
            Invoke(
                Sub()
                    Dim dc = _mainDrawing.Append()
                    dc.DrawRectangle(Nothing, _pen, New System.Windows.Rect(CInt(x), CInt(y), width, height))
                    dc.Close()
                    AddRasterizeOperationToQueue()
                End Sub)
        End Sub

        ''' <summary>
        ''' Fills a rectangle on the screen using the selected Brush.
        ''' </summary>
        ''' <param name="x">The x co-ordinate of the rectangle.</param>
        ''' <param name="y">The y co-ordinate of the rectangle.</param>
        ''' <param name="width">The width of the rectangle.</param>
        ''' <param name="height">The height of the rectangle.</param>
        Public Shared Sub FillRectangle(x As Primitive, y As Primitive, width As Primitive, height As Primitive)
            VerifyAccess()
            Invoke(
                Sub()
                    Dim dc = _mainDrawing.Append()
                    dc.DrawRectangle(_fillBrush, Nothing, New System.Windows.Rect(x, y, width, height))
                    dc.Close()
                    AddRasterizeOperationToQueue()
                End Sub)
        End Sub

        ''' <summary>
        ''' Draws an ellipse on the screen using the selected Pen.
        ''' </summary>
        ''' <param name="x">The x co-ordinate of the ellipse.</param>
        ''' <param name="y">The y co-ordinate of the ellipse.</param>
        ''' <param name="width">The width of the ellipse.</param>
        ''' <param name="height">The height of the ellipse.</param>
        Public Shared Sub DrawEllipse(x As Primitive, y As Primitive, width As Primitive, height As Primitive)
            VerifyAccess()
            Invoke(
                Sub()
                    Dim dc = _mainDrawing.Append()
                    Dim w = CDbl(width) / 2
                    Dim h = CDbl(height) / 2
                    dc.DrawEllipse(
                        Nothing,
                        _pen, New System.Windows.Point(CDbl(x) + w, CDbl(y) + h),
                        w,
                        h
                    )
                    dc.Close()
                    AddRasterizeOperationToQueue()
                End Sub)
        End Sub


        ''' <summary>
        ''' Draws the polygon that is represented by the given points array with the pen of the graphics window.
        ''' </summary>
        ''' <param name="xOffset">The horizontal offest to add to the x-cordinate of each point of the polygon.</param>
        ''' <param name="yOffset">The vertical offest to add to the y-cordinate of each point of the polygon.</param> 
        ''' <param name="xScale">The factor to multiply the polygon width by.</param>
        ''' <param name="yScale">The factor to multiply the polygon height by.</param>
        ''' <param name="pointsArr">An array of points representing the heads of the polygn. Each item in this array is an array containing the x and y of the point.</param>
        Public Shared Sub DrawPolygon(
                   xOffset As Primitive,
                   yOffset As Primitive,
                   xScale As Primitive,
                   yScale As Primitive,
                   pointsArr As Primitive)

            VerifyAccess() ' must be here to create the pen
            CreatePolygon(xOffset, yOffset, xScale, yScale, pointsArr, _pen, Nothing)
        End Sub

        ''' <summary>
        ''' Fills the polygon that is represented by the given points array with the brush of the graphics window.
        ''' </summary>
        ''' <param name="xOffset">The horizontal offest to add to the x-cordinate of each point of the polygon.</param>
        ''' <param name="yOffset">The vertical offest to add to the y-cordinate of each point of the polygon.</param> 
        ''' <param name="xScale">The factor to multiply the polygon width by.</param>
        ''' <param name="yScale">The factor to multiply the polygon height by.</param>
        ''' <param name="pointsArr">An array of points representing the heads of the polygn. Each item in this array is an array containing the x and y of the point.</param>
        Public Shared Sub FillPolygon(
                   xOffset As Primitive,
                   yOffset As Primitive,
                   xScale As Primitive,
                   yScale As Primitive,
                   pointsArr As Primitive)

            VerifyAccess() ' must be here to create the brush
            CreatePolygon(xOffset, yOffset, xScale, yScale, pointsArr, Nothing, _fillBrush)
        End Sub

        Private Shared Sub CreatePolygon(
                   xOffset As Primitive,
                   yOffset As Primitive,
                   xScale As Primitive,
                   yScale As Primitive,
                   pointsArr As Primitive,
                   pen As Media.Pen,
                   brush As Media.Brush)

            GraphicsWindow.Invoke(
                Sub()
                    If pointsArr.IsEmpty OrElse Not pointsArr.IsArray Then Return

                    Dim dc = _mainDrawing.Append()
                    Dim Points As New PointCollection()
                    Dim x1, x2, y1, y2 As Double

                    For Each point In pointsArr.ArrayMap.Values
                        Dim x As Double = point.Items(1)
                        Dim y As Double = point.Items(2)
                        Points.Add(New System.Windows.Point(x, y))

                        If x < x1 Then
                            x1 = x
                        ElseIf x > x2 Then
                            x2 = x
                        End If

                        If y < y1 Then
                            y1 = y
                        ElseIf y > y2 Then
                            y2 = y
                        End If
                    Next

                    Dim polygon As New System.Windows.Shapes.Polygon With {
                        .Points = Points,
                        .Fill = brush
                    }

                    If pen IsNot Nothing Then
                        polygon.Stroke = pen.Brush
                        polygon.StrokeThickness = pen.Thickness
                    End If

                    dc.DrawRectangle(
                        New VisualBrush(polygon), Nothing,
                        New System.Windows.Rect(
                                x1 + xOffset,
                                y1 + yOffset,
                                (x2 - x1) * xScale,
                                (y2 - y1) * yScale
                        )
                    )

                    dc.Close()
                    AddRasterizeOperationToQueue()
                End Sub)
        End Sub


        ''' <summary>
        ''' Fills an ellipse on the screen using the selected Brush.
        ''' </summary>
        ''' <param name="x">The x co-ordinate of the ellipse.</param>
        ''' <param name="y">The y co-ordinate of the ellipse.</param>
        ''' <param name="width">The width of the ellipse.</param>
        ''' <param name="height">The height of the ellipse.</param>
        Public Shared Sub FillEllipse(x As Primitive, y As Primitive, width As Primitive, height As Primitive)
            VerifyAccess()
            Invoke(Sub()
                       Dim drawingContext = _mainDrawing.Append()
                       Dim w = CDbl(width) / 2
                       Dim h = CDbl(height) / 2
                       drawingContext.DrawEllipse(
                                _fillBrush,
                                Nothing,
                                New System.Windows.Point(CDbl(x) + w, CDbl(y) + h),
                                w,
                                h
                            )
                       drawingContext.Close()
                       AddRasterizeOperationToQueue()
                   End Sub)
        End Sub

        ''' <summary>
        ''' Draws a triangle on the screen using the selected pen.
        ''' </summary>
        ''' <param name="x1">The x co-ordinate of the first point.</param>
        ''' <param name="y1">The y co-ordinate of the first point.</param>
        ''' <param name="x2">The x co-ordinate of the second point.</param>
        ''' <param name="y2">The y co-ordinate of the second point.</param>
        ''' <param name="x3">The x co-ordinate of the third point.</param>
        ''' <param name="y3">The y co-ordinate of the third point.</param>
        Public Shared Sub DrawTriangle(
                         x1 As Primitive, y1 As Primitive,
                         x2 As Primitive, y2 As Primitive,
                         x3 As Primitive, y3 As Primitive)

            VerifyAccess()
            Invoke(
                Sub()
                    Dim dc = _mainDrawing.Append()
                    Dim figure As New PathFigure With {
                            .StartPoint = New System.Windows.Point(x1, y1),
                            .IsClosed = True,
                            .Segments = New PathSegmentCollection From {
                                    CType(New LineSegment(New System.Windows.Point(x2, y2), isStroked:=True), PathSegment),
                                    CType(New LineSegment(New System.Windows.Point(x3, y3), isStroked:=True), PathSegment)
                            }
                    }
                    Dim figures As New PathFigureCollection From {figure}
                    Dim triangle As New PathGeometry
                    triangle.Figures = figures
                    triangle.Freeze()
                    dc.DrawGeometry(Nothing, _pen, triangle)
                    dc.Close()
                    AddRasterizeOperationToQueue()
                End Sub)
        End Sub

        ''' <summary>
        ''' Draws and fills a triangle on the screen using the selected brush.
        ''' </summary>
        ''' <param name="x1">The x co-ordinate of the first point.</param>
        ''' <param name="y1">The y co-ordinate of the first point.</param>
        ''' <param name="x2">The x co-ordinate of the second point.</param>
        ''' <param name="y2">The y co-ordinate of the second point.</param>
        ''' <param name="x3">The x co-ordinate of the third point.</param>
        ''' <param name="y3">The y co-ordinate of the third point.</param>
        Public Shared Sub FillTriangle(x1 As Primitive, y1 As Primitive, x2 As Primitive, y2 As Primitive, x3 As Primitive, y3 As Primitive)
            VerifyAccess()
            Invoke(
                Sub()
                    Dim dc As DrawingContext = _mainDrawing.Append()
                    Dim figure As New PathFigure With {
                            .StartPoint = New System.Windows.Point(x1, y1),
                            .IsClosed = True,
                            .Segments = New PathSegmentCollection From {
                                    CType(New LineSegment(New System.Windows.Point(x2, y2), isStroked:=True), PathSegment),
                                    CType(New LineSegment(New System.Windows.Point(x3, y3), isStroked:=True), PathSegment)
                            }
                    }
                    Dim figures As New PathFigureCollection From {figure}
                    Dim triangle As New PathGeometry
                    triangle.Figures = figures
                    triangle.Freeze()
                    dc.DrawGeometry(_fillBrush, Nothing, triangle)
                    dc.Close()
                    AddRasterizeOperationToQueue()
                End Sub)
        End Sub

        ''' <summary>
        ''' Draws a cubic Bezier curve between the given start and end points, and passing through the two given control points.
        ''' </summary>
        ''' <param name="x1">The x co-ordinate of the start point.</param>
        ''' <param name="y1">The y co-ordinate of the start point.</param>
        ''' <param name="x2">The x co-ordinate of the first control point.</param>
        ''' <param name="y2">The y co-ordinate of the first control point.</param>
        ''' <param name="x3">The x co-ordinate of the second control point.</param>
        ''' <param name="y3">The y co-ordinate of the second control point.</param>
        ''' <param name="x4">The x co-ordinate of the end point.</param>
        ''' <param name="y4">The y co-ordinate of the end point.</param>
        Public Shared Sub DrawBezierCurve(
                         x1 As Primitive, y1 As Primitive,
                         x2 As Primitive, y2 As Primitive,
                         x3 As Primitive, y3 As Primitive,
                         x4 As Primitive, y4 As Primitive
                   )

            VerifyAccess()
            Invoke(
                Sub()
                    Dim dc = _mainDrawing.Append()
                    Dim figure As New PathFigure With {
                            .StartPoint = New System.Windows.Point(x1, y1),
                            .IsClosed = False,
                            .Segments = New PathSegmentCollection From {
                                New BezierSegment(
                                        New System.Windows.Point(x2, y2),
                                        New System.Windows.Point(x3, y3),
                                        New System.Windows.Point(x4, y4),
                                        True
                                )
                            }
                    }
                    Dim figures As New PathFigureCollection From {figure}
                    Dim pg As New PathGeometry
                    pg.Figures = figures
                    pg.Freeze()
                    dc.DrawGeometry(Nothing, _pen, pg)
                    dc.Close()
                    AddRasterizeOperationToQueue()
                End Sub)
        End Sub

        ''' <summary>
        ''' Draws a quadratic Bezier curve, from the given start point, passing through the given control point, and ending at the given end point.
        ''' </summary>
        ''' <param name="x1">The x co-ordinate of the start point.</param>
        ''' <param name="y1">The y co-ordinate of the start point.</param>
        ''' <param name="x2">The x co-ordinate of the control point.</param>
        ''' <param name="y2">The y co-ordinate of the control point.</param>
        ''' <param name="x3">The x co-ordinate of the end point.</param>
        ''' <param name="y3">The y co-ordinate of the end point.</param>
        Public Shared Sub DrawQuadraticBezierCurve(
                         x1 As Primitive, y1 As Primitive,
                         x2 As Primitive, y2 As Primitive,
                         x3 As Primitive, y3 As Primitive
                   )

            VerifyAccess()
            Invoke(
                Sub()
                    Dim dc = _mainDrawing.Append()
                    Dim figure As New PathFigure With {
                            .StartPoint = New System.Windows.Point(x1, y1),
                            .IsClosed = False,
                            .Segments = New PathSegmentCollection From {
                                New QuadraticBezierSegment(
                                        New System.Windows.Point(x2, y2),
                                        New System.Windows.Point(x3, y3),
                                        True
                                )
                            }
                    }
                    Dim figures As New PathFigureCollection From {figure}
                    Dim pg As New PathGeometry
                    pg.Figures = figures
                    pg.Freeze()
                    dc.DrawGeometry(Nothing, _pen, pg)
                    dc.Close()
                    AddRasterizeOperationToQueue()
                End Sub)
        End Sub

        ''' <summary>
        ''' Draws an arc from between the given start and end points. 
        ''' This arc is a part of the ellipse that has the given radius and passes through these tow points.
        ''' </summary>
        ''' <param name="x1">The x co-ordinate of the start point.</param>
        ''' <param name="y1">The y co-ordinate of the start point.</param>
        ''' <param name="x2">The x co-ordinate of the end point of the arc.</param>
        ''' <param name="y2">The y co-ordinate of the end point of the arc.</param>
        ''' <param name="xRadius">The horizontal radius of the arc</param>
        ''' <param name="yRadius">The vertical radius of the arc</param>
        ''' <param name="angle">The x-axis rotation of the ellipse</param>
        ''' <param name="isLargArc">Use True if the arc should be greater than 180 degrees, or False otherwise</param>
        ''' <param name="isClockwise">Use True to draw the arc in a positive angle direction, or False otherwise</param>
        Public Shared Sub DrawArc(
                       x1 As Primitive,
                       y1 As Primitive,
                       x2 As Primitive,
                       y2 As Primitive,
                       xRadius As Primitive,
                       yRadius As Primitive,
                       angle As Primitive,
                       isLargArc As Primitive,
                       isClockwise As Primitive
                   )
            VerifyAccess()
            Invoke(
                Sub()
                    Dim dc = _mainDrawing.Append()
                    Dim figure As New PathFigure With {
                            .StartPoint = New System.Windows.Point(x1, y1),
                            .IsClosed = False,
                            .Segments = New PathSegmentCollection From {
                                    New ArcSegment(
                                         New System.Windows.Point(x2, y2),
                                         New System.Windows.Size(xRadius, yRadius),
                                         angle,
                                         isLargArc,
                                         If(isClockwise, SweepDirection.Clockwise, SweepDirection.Counterclockwise),
                                         True
                                    )
                            }
                    }
                    Dim figures As New PathFigureCollection From {figure}
                    Dim pg As New PathGeometry
                    pg.Figures = figures
                    pg.Freeze()
                    dc.DrawGeometry(Nothing, _pen, pg)
                    dc.Close()
                    AddRasterizeOperationToQueue()
                End Sub)
        End Sub

        ''' <summary>
        ''' Draws a line from one point to another.
        ''' </summary>
        ''' <param name="x1">The x co-ordinate of the first point.</param>
        ''' <param name="y1">The y co-ordinate of the first point.</param>
        ''' <param name="x2">The x co-ordinate of the second point.</param>
        ''' <param name="y2">The y co-ordinate of the second point.</param>
        Public Shared Sub DrawLine(x1 As Primitive, y1 As Primitive, x2 As Primitive, y2 As Primitive)
            VerifyAccess()
            Invoke(
                Sub()
                    Dim dc As DrawingContext = _mainDrawing.Append()
                    dc.DrawLine(
                        _pen,
                        New System.Windows.Point(x1, y1),
                        New System.Windows.Point(x2, y2)
                    )
                    dc.Close()
                    AddRasterizeOperationToQueue()
                End Sub)
        End Sub

        ''' <summary>
        ''' Draws a line of text on the screen at the specified location.
        ''' </summary>
        ''' <param name="x">The x co-ordinate of the text start point.</param>
        ''' <param name="y">The y co-ordinate of the text start point.</param>
        ''' <param name="text">The text to draw</param>
        Public Shared Sub DrawText(x As Primitive, y As Primitive, text As Primitive)
            If Not text.IsEmpty Then
                VerifyAccess()
                Invoke(
                    Sub()
                        Dim dc = _mainDrawing.Append()
                        Dim formattedText1 As New FormattedText(
                                text,
                                CultureInfo.CurrentUICulture,
                                FlowDirection.LeftToRight,
                                New Typeface(_fontFamily, _fontStyle, _fontWeight, FontStretches.Normal),
                                _fontSize,
                                _fillBrush
                        )
                        dc.DrawText(formattedText1, New System.Windows.Point(x, y))
                        dc.Close()
                        AddRasterizeOperationToQueue()
                    End Sub)
            End If
        End Sub

        ''' <summary>
        ''' Draws a line of text on the screen at the specified location.
        ''' </summary>
        ''' <param name="x">The x co-ordinate of the text start point.</param>
        ''' <param name="y">The y co-ordinate of the text start point.</param>
        ''' <param name="width">The maximum available width.  This parameter helps define when the text should wrap.</param>
        ''' <param name="text">The text to draw.</param>
        Public Shared Sub DrawBoundText(x As Primitive, y As Primitive, width As Primitive, text As Primitive)
            If Not text.IsEmpty Then
                VerifyAccess()
                Invoke(
                    Sub()
                        Dim dc = _mainDrawing.Append()
                        Dim fontTypeFace = New Typeface(
                            _fontFamily,
                            _fontStyle,
                            _fontWeight,
                            FontStretches.Normal
                        )

                        Dim formattedText = New FormattedText(
                            text,
                            CultureInfo.CurrentUICulture,
                            FlowDirection.LeftToRight,
                            fontTypeFace,
                            _fontSize,
                            _fillBrush
                        ) With {.MaxTextWidth = width}

                        dc.DrawText(
                            formattedText,
                            New System.Windows.Point(x, y)
                        )

                        dc.Close()
                        AddRasterizeOperationToQueue()
                    End Sub)
            End If
        End Sub

        ''' <summary>
        ''' Draws the specified image from memory on to the screen, in the specified size.
        ''' </summary>
        ''' <param name="imageName">The name of the image to draw</param>
        ''' <param name="x">The x co-ordinate of the point to draw the image at.</param>
        ''' <param name="y">The y co-ordinate of the point to draw the image at.</param>
        ''' <param name="width">The width to draw the image.</param>
        ''' <param name="height">The height to draw the image.</param>
        Public Shared Sub DrawResizedImage(imageName As Primitive, x As Primitive, y As Primitive, width As Primitive, height As Primitive)
            If imageName.IsEmpty Then
                Return
            End If
            VerifyAccess()
            Dim image1 As BitmapSource = ImageList.GetBitmap(imageName)
            If image1 IsNot Nothing Then
                Invoke(
                    Sub()
                        Dim dc As DrawingContext = _mainDrawing.Append()
                        dc.DrawImage(image1, New System.Windows.Rect(x, y, width, height))
                        dc.Close()
                        AddRasterizeOperationToQueue()
                    End Sub)
            End If
        End Sub

        ''' <summary>
        ''' Draws the specified image from memory on to the screen.  
        ''' </summary>
        ''' <param name="imageName">The name of the image to draw.</param>
        ''' <param name="x">The x co-ordinate of the point to draw the image at.</param>
        ''' <param name="y">The y co-ordinate of the point to draw the image at.</param>
        Public Shared Sub DrawImage(imageName As Primitive, x As Primitive, y As Primitive)
            If imageName.IsEmpty Then Return

            VerifyAccess()
            Dim image1 = ImageList.GetBitmap(imageName)
            If image1 IsNot Nothing Then
                Invoke(
                    Sub()
                        Dim dc As DrawingContext = _mainDrawing.Append()
                        dc.DrawImage(image1, New System.Windows.Rect(x, y, image1.PixelWidth, image1.PixelHeight))
                        dc.Close()
                        AddRasterizeOperationToQueue()
                    End Sub)
            End If
        End Sub

        Private Shared dispatcher As Threading.Dispatcher = SmallBasicApplication.Dispatcher

        ''' <summary>
        ''' Draws the pixel specified by the x and y co-ordinates using the specified color.
        ''' </summary>
        ''' <param name="x">The x co-ordinate of the pixel.</param>
        ''' <param name="y">The y co-ordinate of the pixel.</param>
        ''' <param name="color">The color of the pixel to set.</param>
        Public Shared Sub SetPixel(x As Primitive, y As Primitive, color As Primitive)
            VerifyAccess()
            dispatcher.Invoke(Threading.DispatcherPriority.Render,
                Sub()
                    Dim dc = _mainDrawing.Append()
                    dc.DrawRectangle(
                        New SolidColorBrush(GetColorFromString(color)),
                        Nothing,
                        New System.Windows.Rect(x, y, 1.0, 1.0)
                    )
                    dc.Close()
                    AddRasterizeOperationToQueue()
                End Sub)
        End Sub

        'Private Shared _solidBrushes As New Dictionary(Of String, SolidColorBrush)
        'Private Shared Function GetSolidBrush(color As Primitive) As Media.Brush
        '    Dim b As SolidColorBrush
        '    If Not _solidBrushes.TryGetValue(color, b) Then
        '        b = New SolidColorBrush(GetColorFromString(color))
        '        _solidBrushes(color) = b
        '    End If
        '    Return b
        'End Function

        ''' <summary>
        ''' Gets the color of the pixel at the specified x and y co-ordinates.
        ''' </summary>
        ''' <param name="x">The x co-ordinate of the pixel.</param>
        ''' <param name="y">The y co-ordinate of the pixel.</param>
        ''' <returns>The color of the pixel.</returns>
        <WinForms.ReturnValueType(VariableType.Color)>
        Public Shared Function GetPixel(x As Primitive, y As Primitive) As Primitive
            VerifyAccess()
            If x < 0 Then x = 0
            If y < 0 Then y = 0
            If x > Width Then x = Width
            If y > Height Then y = Height

            Return InvokeWithReturn(
                Function() As Primitive
                    Rasterize()
                    Dim colorBGR As Byte() = New Byte(3) {}
                    Dim stride = CInt(_renderBitmap.Width * (_renderBitmap.Format.BitsPerPixel + 7) / 8)
                    _renderBitmap.CopyPixels(New Int32Rect(x, y, 1, 1), colorBGR, stride, 0)
                    Return New Primitive($"#{colorBGR(2):X2}{colorBGR(1):X2}{colorBGR(0):X2}")
                End Function)
        End Function

        ''' <summary>
        ''' Gets a valid random color.
        ''' </summary>
        ''' <returns>A valid random color.</returns>
        <WinForms.ReturnValueType(VariableType.Color)>
        Public Shared Function GetRandomColor() As Primitive
            If _random Is Nothing Then
                _random = New Random(Now.Ticks Mod Integer.MaxValue)
            End If

            Return New Primitive($"#{_random.Next(256):X2}{_random.Next(256):X2}{_random.Next(256):X2}")
        End Function

        ''' <summary>
        ''' Constructs a color given the Red, Green and Blue values.
        ''' </summary>
        ''' <param name="red">The red component of the Color (0-255).</param>
        ''' <param name="green">The green component of the color (0-255).</param>
        ''' <param name="blue">The blue component of the color (0-255).</param>
        ''' <returns>
        ''' Returns a color that can be used to set the brush or pen color.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.Color)>
        Public Shared Function GetColorFromRGB(red As Primitive, green As Primitive, blue As Primitive) As Primitive
            Dim num As Integer = Math.Abs(CInt(red) Mod 256)
            Dim num2 As Integer = Math.Abs(CInt(green) Mod 256)
            Dim num3 As Integer = Math.Abs(CInt(blue) Mod 256)
            Return New Primitive($"#{num:X2}{num2:X2}{num3:X2}")
        End Function

        ''' <summary>
        ''' Clears the window.
        ''' </summary>
        Public Shared Sub Clear()
            VerifyAccess()
            If Not _windowCreated Then Return

            Invoke(
                Sub()
                    _mainCanvas.Visibility = Visibility.Hidden
                    If Turtle.IsVisible Then Turtle.Initialize(False)
                    _mainCanvas.Children.Clear()
                    _renderBitmap.Clear()
                    _mainDrawing.Children.Clear()
                    _objectsMap.Clear()
                    _mainCanvas.Visibility = Visibility.Visible
                End Sub)
        End Sub

        ''' <summary>
        ''' Displays a message box to the user.
        ''' Use MsgBox as a shorcut name to show the message box. Ex:
        ''' MsgBox "Hello!"
        ''' </summary>
        ''' <param name="message">The text to be displayed on the message box.</param>
        ''' <param name="title">The title for the message box.</param>
        Public Shared Sub ShowMessage(message As Primitive, title As Primitive)
            VerifyAccess()
            Invoke(Sub() MessageBox.Show(message, title))
        End Sub

        ''' <summary>
        ''' Show the graphics window and gets the form object that repreesents it.
        ''' </summary>
        ''' <returns>
        ''' The graphics window name which is "graphicswindow".
        ''' Note that sVB can deal with this name as a form object, so, you can use it to access the properties and methods of the form that repreesents the graphics window.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.Form)>
        Public Shared Function AsForm() As Primitive
            VerifyAccess()
            Return Controls.GW_NAME
        End Function

        Private Shared Sub CreateWindow(Optional keepValue As Boolean = False)
            SyncLock _syncLock
                Invoke(
                    Sub()
                        If Not keepValue Then _backgroundColor = WinForms.Colors.White
                        _brushColor = WinForms.Colors.CornflowerBlue
                        _fillBrush = Media.Brushes.CornflowerBlue
                        _penColor = WinForms.Colors.Black
                        _penWidth = New Primitive(2)
                        _pen = New Media.Pen(Media.Brushes.Black, 2.0)
                        _fontFamily = New Media.FontFamily("Tahoma")
                        _fullScreen = New Primitive(False)
                        _windowCreated = True
                        Shapes._nameGenerationMap.Clear()
                        Shapes._rotateTransformMap.Clear()
                        Shapes._scaleTransformMap.Clear()
                        Shapes._positionMap.Clear()

                        Dim allowsTransparency = (
                                 _backgroundColor = WinForms.Colors.Transparent OrElse
                                _backgroundColor = WinForms.Colors.None OrElse
                                WinForms.Color.GetTransparency(_backgroundColor) > 0
                        )
                        _window = New Window() With {
                            .Name = Controls.GW_NAME,
                            .Title = "sVB Graphics Window",
                            .AllowsTransparency = allowsTransparency,
                            .WindowStyle = If(allowsTransparency, WindowStyle.None, WindowStyle.ThreeDBorderWindow),
                            .Background = WinForms.Color.GetBrush(_backgroundColor),
                            .Height = 480.0,
                            .Width = 640.0
                        }

                        If Not _isHidden Then
                            _windowVisible = True
                            _window.Show()
                        End If

                        AddHandler _window.SourceInitialized,
                             Sub()
                                 Dim handle1 As IntPtr = CType(PresentationSource.FromVisual(_window), HwndSource).Handle
                                 Dim windowLong As UInteger = NativeHelper.GetWindowLong(handle1, -16)
                                 windowLong = (windowLong And &HFFFEFFFFUI) Or &H20000UI
                                 NativeHelper.SetWindowLong(handle1, -16, windowLong)
                                 Dim lpRect As New Internal.RECT With {
                                     .Left = 0,
                                     .Top = 0,
                                     .Right = _window.Width,
                                     .Bottom = _window.Height
                                 }
                                 NativeHelper.AdjustWindowRect(lpRect, windowLong, bMenu:=False)
                                 _window.Width = lpRect.Right - lpRect.Left
                                 _window.Height = lpRect.Bottom - lpRect.Top
                                 NativeHelper.SetForegroundWindow(handle1)
                             End Sub

                        SetWindowContent()
                        SignupForWindowEvents()
                        _window.UpdateLayout()
                        WinForms.Forms._forms(Controls.GW_NAME) = _window
                    End Sub)
            End SyncLock
        End Sub

        Private Shared Sub SignupForWindowEvents()
            AddHandler _window.SizeChanged, AddressOf WindowSizeChanged
            AddHandler _window.Closing, AddressOf WindowClosing
            AddHandler _window.Closed, AddressOf WinForms.Forms.Form_Closed
            AddHandler _window.MouseDown, Sub() RaiseEvent MouseDown()
            AddHandler _window.MouseUp, Sub() RaiseEvent MouseUp()

            AddHandler _window.MouseMove,
                Sub(sender As Object, e As MouseEventArgs)
                    Dim position As System.Windows.Point = e.GetPosition(_mainCanvas)
                    _mouseX = position.X
                    _mouseY = position.Y
                    RaiseEvent MouseMove()
                End Sub

            AddHandler _window.KeyDown,
                Sub(sender As Object, e As KeyEventArgs)
                    _lastKey = e.Key
                    If e.Key = Key.F11 Then
                        FullScreen = Not CBool(FullScreen)
                    ElseIf FullScreen AndAlso e.Key = Key.Escape Then
                        FullScreen = False
                    End If
                    Invoke(Sub() RaiseEvent KeyDown())
                End Sub

            AddHandler _window.TextInput,
                Sub(sender As Object, e As TextCompositionEventArgs)
                    _lastText = e.Text
                    Invoke(Sub() RaiseEvent TextInput())
                End Sub

            AddHandler _window.KeyUp,
                Sub(sender As Object, e As KeyEventArgs)
                    _lastKey = e.Key
                    Invoke(Sub() RaiseEvent KeyUp())
                End Sub

        End Sub

        Private Shared Sub SetWindowContent()
            _mainCanvas = New Canvas With {
                .HorizontalAlignment = HorizontalAlignment.Stretch,
                .VerticalAlignment = VerticalAlignment.Stretch
            }
            _renderBitmap = New RenderTargetBitmap(
                _window.Width + 120,
                _window.Height + 120,
                96.0, 96.0,
                PixelFormats.Default
            )
            _bitmapContainer = New System.Windows.Controls.Image With {
                .Source = _renderBitmap,
                .Stretch = Stretch.None
            }
            _mainDrawing = New DrawingGroup
            _visualContainer = New VisualContainer(_mainDrawing)
            Dim grid1 As New Grid
            grid1.Children.Add(_bitmapContainer)
            grid1.Children.Add(_visualContainer)
            grid1.Children.Add(_mainCanvas)
            _window.Content = grid1
        End Sub

        Private Shared Sub WindowClosing(sender As Object, e As CancelEventArgs)
            Helper.RunLater(_window,
                Sub()
                    If WinForms.Event.Handled Then
                        e.Cancel = True
                        WinForms.Event.Handled = False
                        Return
                    End If

                    _windowCreated = False
                    If WinForms.Forms._forms.Count = 0 AndAlso Not TextWindow._windowVisible Then
                        Program.End()
                    End If

                    Turtle.Initialize()
                    _mainCanvas.Children.Clear()
                    _renderBitmap.Clear()
                    _mainDrawing.Children.Clear()
                    _objectsMap.Clear()
                End Sub, 500)
        End Sub

        Private Shared Sub WindowSizeChanged(sender As Object, e As SizeChangedEventArgs)
            If e.NewSize.Width > _renderBitmap.Width OrElse e.NewSize.Height > _renderBitmap.Height Then
                Dim renderBitmap = _renderBitmap
                _renderBitmap = New RenderTargetBitmap(
                    e.NewSize.Width + 120,
                    e.NewSize.Height + 120,
                    96.0,
                    96.0,
                    PixelFormats.Default
                 )
                Dim drawingVisual1 As New DrawingVisual
                Dim dc = drawingVisual1.RenderOpen()
                dc.DrawImage(renderBitmap, New System.Windows.Rect(0.0, 0.0, renderBitmap.Width, renderBitmap.Height))
                dc.Close()
                _renderBitmap.Render(drawingVisual1)
                _bitmapContainer.Width = _renderBitmap.Width
                _bitmapContainer.Height = _renderBitmap.Height
                _bitmapContainer.Source = _renderBitmap
                _bitmapContainer.Stretch = Stretch.None
            End If
        End Sub

        Friend Shared Sub VerifyAccess(Optional keepValue As Boolean = False)
            If Not CBool(_AutoShow) Then
                If _windowCreated Then BeginInvoke(Sub() Return)
            ElseIf _windowCreated Then
                If _isHidden Then
                    Invoke(Sub()
                               _window.Show()
                               _windowVisible = True
                           End Sub)
                End If
            Else
                CreateWindow(keepValue)
                Invoke(Sub() _window.WindowState = WindowState.Maximized)
            End If
        End Sub

        Friend Shared Sub BeginInvoke(invokeDelegate As InvokeHelper)
            SmallBasicApplication.BeginInvoke(invokeDelegate)
        End Sub

        Friend Shared Sub Invoke(invokeDelegate As InvokeHelper)
            dispatcher.Invoke(Threading.DispatcherPriority.Render, invokeDelegate)
        End Sub

        Friend Shared Function InvokeWithReturn(invokeDelegate As InvokeHelperWithReturn) As Primitive
            Return SmallBasicApplication.InvokeWithReturn(invokeDelegate)
        End Function

        Friend Shared Sub AddShape(
                   name As String,
                   shape As FrameworkElement,
                   Optional addToCanvas As Boolean = True)

            VerifyAccess()
            If name.StartsWith(Controls.GW_NAME) Then
                shape.Name = name.Substring(Controls.GW_NAME.AsString().Length + 1)
            Else
                shape.Name = name
            End If
            _objectsMap(name) = shape
            If addToCanvas Then _mainCanvas.Children.Add(shape)
        End Sub

        Friend Shared Sub AddControl(
                          name As String,
                          control As Control,
                          Optional addToCanvas As Boolean = True
                   )

            VerifyAccess()
            control.Foreground = _fillBrush
            control.FontFamily = _fontFamily
            control.FontStyle = _fontStyle
            control.FontSize = _fontSize
            control.FontWeight = _fontWeight
            AddShape(name, control, addToCanvas)
            WinForms.Forms._controls(name) = control
        End Sub

        Friend Shared Sub RemoveShape(name As Primitive)
            Dim shape As UIElement = Nothing
            If _objectsMap.TryGetValue(name, shape) Then
                _objectsMap.Remove(name)
                Invoke(Sub() _mainCanvas.Children.Remove(shape))
            End If
        End Sub

        Friend Shared Function ContainsShape(name As String) As Boolean
            Return _objectsMap.ContainsKey(name)
        End Function

        Friend Shared Sub AddRasterizeOperationToQueue()
            If _mainDrawing.Children.Count > 100 Then
                Rasterize()
            End If
        End Sub

        Private Shared Sub Rasterize()
            If _mainDrawing.Children.Count <> 0 Then
                _renderBitmap.Render(_visualContainer)
                _mainDrawing.Children.Clear()
            End If
        End Sub

        Friend Shared Sub DoubleAnimateProperty(
                   obj As IAnimatable,
                   [property] As DependencyProperty,
                   [end] As Double,
                   duration As Double)

            Dim dpo = CType(obj, DependencyObject)
            Dim start = CDbl(dpo.GetValue([property]))
            If Double.IsNaN(start) Then start = 0.0
            If start = [end] Then Return

            Dim animation As New DoubleAnimation(
                start,
                [end],
                TimeSpan.FromMilliseconds(duration)
            )

            animation.FillBehavior = FillBehavior.HoldEnd
            animation.DecelerationRatio = 0.2
            obj.BeginAnimation([property], animation)
        End Sub

        Friend Shared Function GetColorFromString(color As String) As Media.Color
            Try
                Return Media.ColorConverter.ConvertFromString(color)
            Catch
            End Try

            Return Colors.Black
        End Function

        Friend Shared Function GetStringFromColor(color1 As Media.Color) As String
            Dim c As System.Drawing.Color = System.Drawing.Color.FromArgb(color1.A, color1.R, color1.G, color1.B)
            Return ColorTranslator.ToHtml(c)
        End Function

        Friend Shared Sub SetCursor(cursor1 As Cursor)
            VerifyAccess()
            Invoke(Sub() _window.Cursor = cursor1)
        End Sub


    End Class
End Namespace
