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
Imports SmallBasicLibrary.Microsoft.SmallBasic.Library.Internal

Namespace Microsoft.SmallBasic.Library
    ''' <summary>
    ''' The GraphicsWindow provides graphics related input and output functionality.  For example, using this class, it is possible to draw and fill circles and rectangles.
    ''' </summary>
    <SmallBasicType>
    Public NotInheritable Class GraphicsWindow
        Private Class VisualContainer
            Inherits FrameworkElement
            Private _drawing As Drawing

            Public Sub New(drawing1 As Drawing)
                _drawing = drawing1
            End Sub
            Protected Overrides Sub OnRender(drawingContext1 As DrawingContext)
                drawingContext1.DrawDrawing(_drawing)
            End Sub
        End Class

        Private Const _defaultFontName As String = "Tahoma"
        Private Const _step As Integer = 20
        Friend Shared _windowVisible As Boolean = False
        Friend Shared _windowCreated As Boolean = False
        Private Shared _window As Window
        Friend Shared _pen As Windows.Media.Pen
        Friend Shared _fillBrush As SolidColorBrush = System.Windows.Media.Brushes.SlateBlue
        Friend Shared _fontFamily As Windows.Media.FontFamily
        Friend Shared _fontSize As Double = 12.0
        Friend Shared _fontStyle As Windows.FontStyle = FontStyles.Normal
        Friend Shared _fontWeight As FontWeight = FontWeights.Bold
        Private Shared _mouseX As Double
        Private Shared _mouseY As Double
        Private Shared _lastKey As Key
        Private Shared _lastText As String
        Private Shared _random As Random
        Private Shared _backgroundBrush As SolidColorBrush
        Friend Shared _objectsMap As New Dictionary(Of String, UIElement)
        Private Shared _scaleTransformMap As New Dictionary(Of String, ScaleTransform)
        Private Shared _mainCanvas As Canvas
        Private Shared _mainDrawing As DrawingGroup
        Private Shared _visualContainer As VisualContainer
        Private Shared _renderBitmap As RenderTargetBitmap
        Private Shared _bitmapContainer As Windows.Controls.Image
        Private Shared _isHidden As Boolean
        Private Shared _syncLock As New Object
        ''' <summary>
        ''' Gets or sets the Background color of the Graphics Window.
        ''' </summary>
        Public Shared Property BackgroundColor As Primitive
            Get
                VerifyAccess()
                Return CStr(InvokeWithReturn(Function() GetStringFromColor(_backgroundBrush.Color)))
            End Get

            Set(Value As Primitive)
                VerifyAccess()
                BeginInvoke(Sub()
                                _backgroundBrush = New SolidColorBrush(GetColorFromString(Value))
                                _window.Background = _backgroundBrush
                            End Sub)
            End Set
        End Property
        ''' <summary>
        ''' Gets or sets the brush color to be used to fill shapes drawn on the Graphics Window.
        ''' </summary>
        Public Shared Property BrushColor As Primitive
            Get
                VerifyAccess()
                Return CStr(InvokeWithReturn(Function() GetStringFromColor(_fillBrush.Color)))
            End Get

            Set(Value As Primitive)
                VerifyAccess()
                BeginInvoke(Sub()
                                _fillBrush = New SolidColorBrush(GetColorFromString(Value))
                                _fillBrush.Freeze()
                            End Sub)
            End Set
        End Property
        ''' <summary>
        ''' Specifies whether or not the Graphics Window can be resized by the user.
        ''' </summary>
        Public Shared Property CanResize As Primitive
            Get
                VerifyAccess()
                Return CBool(InvokeWithReturn(Function() _window.ResizeMode = ResizeMode.CanResize))
            End Get

            Set(Value As Primitive)
                VerifyAccess()
                BeginInvoke(Sub()
                                If CBool(Value) Then
                                    _window.ResizeMode = ResizeMode.CanResize
                                Else
                                    _window.ResizeMode = ResizeMode.NoResize
                                End If
                            End Sub)
            End Set
        End Property
        ''' <summary>
        ''' Gets or sets the width of the pen used to draw shapes on the Graphics Window.
        ''' </summary>
        Public Shared Property PenWidth As Primitive
            Get
                VerifyAccess()
                Return CDbl(InvokeWithReturn(Function() If((_pen IsNot Nothing), (CObj(_pen.Thickness)), (CObj(2.0)))))
            End Get

            Set(Value As Primitive)
                VerifyAccess()
                BeginInvoke(Sub()
                                _pen = New Windows.Media.Pen(_pen.Brush, Value)
                                _pen.Freeze()
                            End Sub)
            End Set
        End Property
        ''' <summary>
        ''' Gets or sets the color of the pen used to draw shapes on the Graphics Window.
        ''' </summary>
        Public Shared Property PenColor As Primitive
            Get
                VerifyAccess()
                Return CStr(InvokeWithReturn(Function() If((_pen IsNot Nothing), GetStringFromColor(CType(_pen.Brush, SolidColorBrush).Color), "Black")))
            End Get

            Set(Value As Primitive)
                VerifyAccess()
                BeginInvoke(Sub()
                                Dim colorFromString As Windows.Media.Color = GetColorFromString(Value)
                                _pen = New Windows.Media.Pen(New SolidColorBrush(colorFromString), _pen.Thickness)
                                _pen.Freeze()
                            End Sub)
            End Set
        End Property
        ''' <summary>
        ''' Gets or sets the Font Name to be used when drawing text on the Graphics Window.
        ''' </summary>
        Public Shared Property FontName As Primitive
            Get
                Return CStr(InvokeWithReturn(Function() If((_fontFamily IsNot Nothing), _fontFamily.Source, "Tahoma")))
            End Get

            Set(Value As Primitive)
                BeginInvoke(Sub()
                                _fontFamily = New Windows.Media.FontFamily(Value)
                            End Sub)
            End Set
        End Property
        ''' <summary>
        ''' Gets or sets the Font Size to be used when drawing text on the Graphics Window.
        ''' </summary>
        Public Shared Property FontSize As Primitive
            Get
                Return New Primitive(_fontSize)
            End Get

            Set(Value As Primitive)
                BeginInvoke(Sub()
                                _fontSize = Value
                            End Sub)
            End Set
        End Property
        ''' <summary>
        ''' Gets or sets whether or not the font to be used when drawing text on the Graphics Window, is bold.
        ''' </summary>
        Public Shared Property FontBold As Primitive
            Get
                Return _fontWeight = FontWeights.Bold
            End Get

            Set(Value As Primitive)
                BeginInvoke(Sub()
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
        Public Shared Property FontItalic As Primitive
            Get
                Return _fontStyle = FontStyles.Italic
            End Get

            Set(Value As Primitive)
                BeginInvoke(Sub()
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
        Public Shared Property Title As Primitive
            Get
                VerifyAccess()
                Return CStr(InvokeWithReturn(Function() _window.Title))
            End Get

            Set(Value As Primitive)
                VerifyAccess()
                BeginInvoke(Sub()
                                _window.Title = Value
                            End Sub)
            End Set
        End Property
        ''' <summary>
        ''' Gets or sets the Height of the graphics window.
        ''' </summary>
        Public Shared Property Height As Primitive
            Get
                VerifyAccess()
                Return CDbl(InvokeWithReturn(Function() _mainCanvas.ActualHeight))
            End Get

            Set(Value As Primitive)
                VerifyAccess()
                BeginInvoke(Sub()
                                _window.Height = Value + (_window.ActualHeight - _mainCanvas.ActualHeight)
                            End Sub)
            End Set
        End Property
        ''' <summary>
        ''' Gets or sets the Width of the graphics window.
        ''' </summary>
        Public Shared Property Width As Primitive
            Get
                VerifyAccess()
                Return CDbl(InvokeWithReturn(Function() _mainCanvas.ActualWidth))
            End Get

            Set(Value As Primitive)
                VerifyAccess()
                BeginInvoke(Sub()
                                _window.Width = Value + (_window.ActualWidth - _mainCanvas.ActualWidth)
                            End Sub)
            End Set
        End Property
        ''' <summary>
        ''' Gets or sets the Left Position of the graphics window.
        ''' </summary>
        Public Shared Property Left As Primitive
            Get
                VerifyAccess()
                Return CDbl(InvokeWithReturn(Function() _window.Left))
            End Get

            Set(Value As Primitive)
                VerifyAccess()
                BeginInvoke(Sub()
                                _window.Left = Value
                            End Sub)
            End Set
        End Property
        ''' <summary>
        ''' Gets or sets the Top Position of the graphics window.
        ''' </summary>
        Public Shared Property Top As Primitive
            Get
                VerifyAccess()
                Return CDbl(InvokeWithReturn(Function() _window.Top))
            End Get

            Set(Value As Primitive)
                VerifyAccess()
                BeginInvoke(Sub()
                                _window.Top = Value
                            End Sub)
            End Set
        End Property
        ''' <summary>
        ''' Gets the last key that was pressed or released.
        ''' </summary>
        Public Shared ReadOnly Property LastKey As Primitive
            Get
                Return _lastKey.ToString()
            End Get
        End Property
        ''' <summary>
        ''' Gets the last text that was entered on the Graphics Window.
        ''' </summary>
        Public Shared ReadOnly Property LastText As Primitive
            Get
                Return _lastText.ToString()
            End Get
        End Property
        ''' <summary>
        ''' Gets the x-position of the mouse relative to the Graphics Window.
        ''' </summary>
        Public Shared ReadOnly Property MouseX As Primitive
            Get
                Return _mouseX
            End Get
        End Property
        ''' <summary>
        ''' Gets the y-position of the mouse relative to the Graphics Window.
        ''' </summary>
        Public Shared ReadOnly Property MouseY As Primitive
            Get
                Return _mouseY
            End Get
        End Property

        Private Shared Events As New EventHandlerList

        ''' <summary>
        ''' Raises an event when a key is pressed down on the keyboard.
        ''' </summary>
        Public Shared Custom Event KeyDown As SmallBasicCallback
            AddHandler(Value As SmallBasicCallback)
                VerifyAccess()
                Dim Key = NameOf(KeyDown)
                Dim h = TryCast(Events(Key), SmallBasicCallback)
                If h IsNot Nothing Then Events.RemoveHandler(Key, h)
                Events.AddHandler(Key, Value)
            End AddHandler

            RemoveHandler(Value As SmallBasicCallback)
                Events.RemoveHandler(NameOf(KeyDown), Value)
            End RemoveHandler

            RaiseEvent()
                Dim h = TryCast(Events(NameOf(KeyDown)), SmallBasicCallback)
                If h IsNot Nothing Then h.Invoke()
            End RaiseEvent
        End Event

        ''' <summary>
        ''' Raises an event when a key is released on the keyboard.
        ''' </summary>
        Public Shared Custom Event KeyUp As SmallBasicCallback
            AddHandler(Value As SmallBasicCallback)
                VerifyAccess()
                Dim Key = NameOf(KeyUp)
                Dim h = TryCast(Events(Key), SmallBasicCallback)
                If h IsNot Nothing Then Events.RemoveHandler(Key, h)
                Events.AddHandler(Key, Value)
            End AddHandler

            RemoveHandler(Value As SmallBasicCallback)
                Events.RemoveHandler(NameOf(KeyUp), Value)
            End RemoveHandler

            RaiseEvent()
                Dim h = TryCast(Events(NameOf(KeyUp)), SmallBasicCallback)
                If h IsNot Nothing Then h.Invoke()
            End RaiseEvent
        End Event

        ''' <summary>
        ''' Raises an event when the mouse button is clicked down.
        ''' </summary>
        Public Shared Custom Event MouseDown As SmallBasicCallback
            AddHandler(Value As SmallBasicCallback)
                VerifyAccess()
                Dim Key = NameOf(MouseDown)
                Dim h = TryCast(Events(Key), SmallBasicCallback)
                If h IsNot Nothing Then Events.RemoveHandler(Key, h)
                Events.AddHandler(Key, Value)
            End AddHandler

            RemoveHandler(Value As SmallBasicCallback)
                Events.RemoveHandler(NameOf(MouseDown), Value)
            End RemoveHandler

            RaiseEvent()
                Dim h = TryCast(Events(NameOf(MouseDown)), SmallBasicCallback)
                If h IsNot Nothing Then h.Invoke()
            End RaiseEvent

        End Event

        ''' <summary>
        ''' Raises an event when the mouse button is released.
        ''' </summary>
        Public Shared Custom Event MouseUp As SmallBasicCallback
            AddHandler(Value As SmallBasicCallback)
                VerifyAccess()
                Dim Key = NameOf(MouseUp)
                Dim h = TryCast(Events(Key), SmallBasicCallback)
                If h IsNot Nothing Then Events.RemoveHandler(Key, h)
                Events.AddHandler(Key, Value)
            End AddHandler

            RemoveHandler(Value As SmallBasicCallback)
                Events.RemoveHandler(NameOf(MouseUp), Value)
            End RemoveHandler

            RaiseEvent()
                Dim h = TryCast(Events(NameOf(MouseUp)), SmallBasicCallback)
                If h IsNot Nothing Then h.Invoke()
            End RaiseEvent

        End Event

        ''' <summary>
        ''' Raises an event when the mouse is moved around.
        ''' </summary>
        Public Shared Custom Event MouseMove As SmallBasicCallback
            AddHandler(Value As SmallBasicCallback)
                VerifyAccess()
                Dim Key = NameOf(MouseMove)
                Dim h = TryCast(Events(Key), SmallBasicCallback)
                If h IsNot Nothing Then Events.RemoveHandler(Key, h)
                Events.AddHandler(Key, Value)
            End AddHandler

            RemoveHandler(Value As SmallBasicCallback)
                Events.RemoveHandler(NameOf(MouseMove), Value)
            End RemoveHandler

            RaiseEvent()
                Dim h = TryCast(Events(NameOf(MouseMove)), SmallBasicCallback)
                If h IsNot Nothing Then h.Invoke()
            End RaiseEvent


        End Event

        ''' <summary>
        ''' Raises an event when text is entered on the GraphicsWindow.
        ''' </summary>
        Public Shared Custom Event TextInput As SmallBasicCallback
            AddHandler(Value As SmallBasicCallback)
                VerifyAccess()
                Dim Key = NameOf(TextInput)
                Dim h = TryCast(Events(Key), SmallBasicCallback)
                If h IsNot Nothing Then Events.RemoveHandler(Key, h)
                Events.AddHandler(Key, Value)
            End AddHandler

            RemoveHandler(Value As SmallBasicCallback)
                Events.RemoveHandler(NameOf(TextInput), Value)
            End RemoveHandler

            RaiseEvent()
                Dim h = TryCast(Events(NameOf(TextInput)), SmallBasicCallback)
                If h IsNot Nothing Then h.Invoke()
            End RaiseEvent


        End Event

        ''' <summary>
        ''' Shows the Graphics window to enable interactions with it.
        ''' </summary>
        Public Shared Sub Show()
            If Not _windowCreated Then
                CreateWindow()
            End If
            _isHidden = False
            If Not _windowVisible Then
                Invoke(Sub()
                           _window.Show()
                       End Sub)
                _windowVisible = True
            End If
        End Sub
        ''' <summary>
        ''' Hides the Graphics window.
        ''' </summary>
        Public Shared Sub Hide()
            _isHidden = True
            If _windowVisible Then
                Invoke(Sub()
                           _window.Hide()
                       End Sub)
                _windowVisible = False
            End If
        End Sub
        ''' <summary>
        ''' Draws a rectangle on the screen using the selected Pen.
        ''' </summary>
        ''' <param name="x">
        ''' The x co-ordinate of the rectangle.
        ''' </param>
        ''' <param name="y">
        ''' The y co-ordinate of the rectangle.
        ''' </param>
        ''' <param name="width">
        ''' The width of the rectangle.
        ''' </param>
        ''' <param name="height">
        ''' The height of the rectangle.
        ''' </param>
        Public Shared Sub DrawRectangle(x As Primitive, y As Primitive, width1 As Primitive, height1 As Primitive)
            VerifyAccess()
            BeginInvoke(
                Sub()
                    Dim drawingContext1 As DrawingContext = _mainDrawing.Append()
                    drawingContext1.DrawRectangle(Nothing, _pen, New Windows.Rect(CInt(x), CInt(y), width1, height1))
                    drawingContext1.Close()
                    AddRasterizeOperationToQueue()
                End Sub)
        End Sub
        ''' <summary>
        ''' Fills a rectangle on the screen using the selected Brush.
        ''' </summary>
        ''' <param name="x">
        ''' The x co-ordinate of the rectangle.
        ''' </param>
        ''' <param name="y">
        ''' The y co-ordinate of the rectangle.
        ''' </param>
        ''' <param name="width">
        ''' The width of the rectangle.
        ''' </param>
        ''' <param name="height">
        ''' The height of the rectangle.
        ''' </param>
        Public Shared Sub FillRectangle(x As Primitive, y As Primitive, width1 As Primitive, height1 As Primitive)
            VerifyAccess()
            BeginInvoke(Sub()
                            Dim drawingContext1 As DrawingContext = _mainDrawing.Append()
                            drawingContext1.DrawRectangle(_fillBrush, Nothing, New Windows.Rect(x, y, width1, height1))
                            drawingContext1.Close()
                            AddRasterizeOperationToQueue()
                        End Sub)
        End Sub
        ''' <summary>
        ''' Draws an ellipse on the screen using the selected Pen.
        ''' </summary>
        ''' <param name="x">
        ''' The x co-ordinate of the ellipse.
        ''' </param>
        ''' <param name="y">
        ''' The y co-ordinate of the ellipse.
        ''' </param>
        ''' <param name="width">
        ''' The width of the ellipse.
        ''' </param>
        ''' <param name="height">
        ''' The height of the ellipse.
        ''' </param>
        Public Shared Sub DrawEllipse(x As Primitive, y As Primitive, width1 As Primitive, height1 As Primitive)
            VerifyAccess()
            BeginInvoke(Sub()
                            Dim drawingContext1 As DrawingContext = _mainDrawing.Append()
                            drawingContext1.DrawEllipse(Nothing, _pen, New Windows.Point(CDbl(x) + CDbl(width1) / 2, CDbl(y) + CDbl(height1) / 2), CDbl(width1) / 2, CDbl(height1) / 2)
                            drawingContext1.Close()
                            AddRasterizeOperationToQueue()
                        End Sub)
        End Sub
        ''' <summary>
        ''' Fills an ellipse on the screen using the selected Brush.
        ''' </summary>
        ''' <param name="x">
        ''' The x co-ordinate of the ellipse.
        ''' </param>
        ''' <param name="y">
        ''' The y co-ordinate of the ellipse.
        ''' </param>
        ''' <param name="width">
        ''' The width of the ellipse.
        ''' </param>
        ''' <param name="height">
        ''' The height of the ellipse.
        ''' </param>
        Public Shared Sub FillEllipse(x As Primitive, y As Primitive, width1 As Primitive, height1 As Primitive)
            VerifyAccess()
            BeginInvoke(Sub()
                            Dim drawingContext1 As DrawingContext = _mainDrawing.Append()
                            drawingContext1.DrawEllipse(_fillBrush, Nothing, New Windows.Point(CDbl(x) + CDbl(width1 / 2), CDbl(y) + CDbl(height1 / 2)), width1 / 2, height1 / 2)
                            drawingContext1.Close()
                            AddRasterizeOperationToQueue()
                        End Sub)
        End Sub
        ''' <summary>
        ''' Draws a triangle on the screen using the selected pen.
        ''' </summary>
        ''' <param name="x1">
        ''' The x co-ordinate of the first point.
        ''' </param>
        ''' <param name="y1">
        ''' The y co-ordinate of the first point.
        ''' </param>
        ''' <param name="x2">
        ''' The x co-ordinate of the second point.
        ''' </param>
        ''' <param name="y2">
        ''' The y co-ordinate of the second point.
        ''' </param>
        ''' <param name="x3">
        ''' The x co-ordinate of the third point.
        ''' </param>
        ''' <param name="y3">
        ''' The y co-ordinate of the third point.
        ''' </param>
        Public Shared Sub DrawTriangle(x1 As Primitive, y1 As Primitive, x2 As Primitive, y2 As Primitive, x3 As Primitive, y3 As Primitive)
            VerifyAccess()
            BeginInvoke(Sub()
                            Dim drawingContext1 As DrawingContext = _mainDrawing.Append()
                            Dim value As New PathFigure With
                            {
                            .StartPoint = New Windows.Point(x1, y1),
                            .IsClosed = True,
                            .Segments = New System.Windows.Media.PathSegmentCollection From {
                                    CType(New LineSegment(New Windows.Point(x2, y2), isStroked:=True), PathSegment),
                                    CType(New LineSegment(New Windows.Point(x3, y3), isStroked:=True), PathSegment)
                            }
                            }
                            Dim figures As New PathFigureCollection From {
                                value
                            }
                            Dim pathGeometry1 As New PathGeometry
                            pathGeometry1.Figures = figures
                            pathGeometry1.Freeze()
                            drawingContext1.DrawGeometry(Nothing, _pen, pathGeometry1)
                            drawingContext1.Close()
                            AddRasterizeOperationToQueue()
                        End Sub)
        End Sub
        ''' <summary>
        ''' Draws and fills a triangle on the screen using the selected brush.
        ''' </summary>
        ''' <param name="x1">
        ''' The x co-ordinate of the first point.
        ''' </param>
        ''' <param name="y1">
        ''' The y co-ordinate of the first point.
        ''' </param>
        ''' <param name="x2">
        ''' The x co-ordinate of the second point.
        ''' </param>
        ''' <param name="y2">
        ''' The y co-ordinate of the second point.
        ''' </param>
        ''' <param name="x3">
        ''' The x co-ordinate of the third point.
        ''' </param>
        ''' <param name="y3">
        ''' The y co-ordinate of the third point.
        ''' </param>
        Public Shared Sub FillTriangle(x1 As Primitive, y1 As Primitive, x2 As Primitive, y2 As Primitive, x3 As Primitive, y3 As Primitive)
            VerifyAccess()
            BeginInvoke(Sub()
                            Dim drawingContext1 As DrawingContext = _mainDrawing.Append()
                            Dim value As New PathFigure With
                            {
                            .StartPoint = New Windows.Point(x1, y1),
                            .IsClosed = True,
                            .Segments = New System.Windows.Media.PathSegmentCollection From {
                                    CType(New LineSegment(New Windows.Point(x2, y2), isStroked:=True), PathSegment),
                                    CType(New LineSegment(New Windows.Point(x3, y3), isStroked:=True), PathSegment)
                            }
                            }
                            Dim figures As New PathFigureCollection From {
                                value
                            }
                            Dim pathGeometry1 As New PathGeometry
                            pathGeometry1.Figures = figures
                            pathGeometry1.Freeze()
                            drawingContext1.DrawGeometry(_fillBrush, Nothing, pathGeometry1)
                            drawingContext1.Close()
                            AddRasterizeOperationToQueue()
                        End Sub)
        End Sub
        ''' <summary>
        ''' Draws a line from one point to another.
        ''' </summary>
        ''' <param name="x1">
        ''' The x co-ordinate of the first point.
        ''' </param>
        ''' <param name="y1">
        ''' The y co-ordinate of the first point.
        ''' </param>
        ''' <param name="x2">
        ''' The x co-ordinate of the second point.
        ''' </param>
        ''' <param name="y2">
        ''' The y co-ordinate of the second point.
        ''' </param>
        Public Shared Sub DrawLine(x1 As Primitive, y1 As Primitive, x2 As Primitive, y2 As Primitive)
            VerifyAccess()
            BeginInvoke(Sub()
                            Dim drawingContext1 As DrawingContext = _mainDrawing.Append()
                            drawingContext1.DrawLine(
                                  _pen,
                                 New Windows.Point(x1, y1),
                                 New Windows.Point(x2, y2)
                             )
                            drawingContext1.Close()
                            AddRasterizeOperationToQueue()
                        End Sub)
        End Sub
        ''' <summary>
        ''' Draws a line of text on the screen at the specified location.
        ''' </summary>
        ''' <param name="x">
        ''' The x co-ordinate of the text start point.
        ''' </param>
        ''' <param name="y">
        ''' The y co-ordinate of the text start point.
        ''' </param>
        ''' <param name="text">
        ''' The text to draw
        ''' </param>
        Public Shared Sub DrawText(x As Primitive, y As Primitive, text As Primitive)
            If Not text.IsEmpty Then
                VerifyAccess()
                BeginInvoke(Sub()
                                Dim drawingContext1 As DrawingContext = _mainDrawing.Append()
                                Dim formattedText1 As New FormattedText(
                                        text,
                                        CultureInfo.CurrentUICulture,
                                        FlowDirection.LeftToRight,
                                        New Typeface(_fontFamily, _fontStyle, _fontWeight, FontStretches.Normal),
                                        _fontSize,
                                        _fillBrush
                                 )
                                drawingContext1.DrawText(formattedText1, New Windows.Point(x, y))
                                drawingContext1.Close()
                                AddRasterizeOperationToQueue()
                            End Sub)
            End If
        End Sub
        ''' <summary>
        ''' Draws a line of text on the screen at the specified location.
        ''' </summary>
        ''' <param name="x">
        ''' The x co-ordinate of the text start point.
        ''' </param>
        ''' <param name="y">
        ''' The y co-ordinate of the text start point.
        ''' </param>
        ''' <param name="width">
        ''' The maximum available width.  This parameter helps define when the text should wrap.
        ''' </param>
        ''' <param name="text">
        ''' The text to draw.
        ''' </param>
        Public Shared Sub DrawBoundText(x As Primitive, y As Primitive, width1 As Primitive, text As Primitive)
            If Not text.IsEmpty Then
                VerifyAccess()
                BeginInvoke(Sub()
                                Dim drawingContext1 As DrawingContext = _mainDrawing.Append()
                                drawingContext1.DrawText(
                                        New FormattedText(text, CultureInfo.CurrentUICulture, FlowDirection.LeftToRight, New Typeface(_fontFamily, _fontStyle, _fontWeight, FontStretches.Normal), _fontSize, _fillBrush) With {.MaxTextWidth = width1},
                                        New Windows.Point(x, y))
                                drawingContext1.Close()
                                AddRasterizeOperationToQueue()
                            End Sub)
            End If
        End Sub
        ''' <summary>
        ''' Draws the specified image from memory on to the screen, in the specified size.
        ''' </summary>
        ''' <param name="imageName">
        ''' The name of the image to draw
        ''' </param>
        ''' <param name="x">
        ''' The x co-ordinate of the point to draw the image at.
        ''' </param>
        ''' <param name="y">
        ''' The y co-ordinate of the point to draw the image at.
        ''' </param>
        ''' <param name="width">
        ''' The width to draw the image.
        ''' </param>
        ''' <param name="height">
        ''' The height to draw the image.
        ''' </param>
        Public Shared Sub DrawResizedImage(imageName As Primitive, x As Primitive, y As Primitive, width1 As Primitive, height1 As Primitive)
            If imageName.IsEmpty Then
                Return
            End If
            VerifyAccess()
            Dim image1 As BitmapSource = ImageList.GetBitmap(imageName)
            If image1 IsNot Nothing Then
                BeginInvoke(
                    Sub()
                        Dim drawingContext1 As DrawingContext = _mainDrawing.Append()
                        drawingContext1.DrawImage(image1, New Windows.Rect(x, y, width1, height1))
                        drawingContext1.Close()
                        AddRasterizeOperationToQueue()
                    End Sub)
            End If
        End Sub
        ''' <summary>
        ''' Draws the specified image from memory on to the screen.  
        ''' </summary>
        ''' <param name="imageName">
        ''' The name of the image to draw.
        ''' </param>
        ''' <param name="x">
        ''' The x co-ordinate of the point to draw the image at.
        ''' </param>
        ''' <param name="y">
        ''' The y co-ordinate of the point to draw the image at.
        ''' </param>
        Public Shared Sub DrawImage(imageName As Primitive, x As Primitive, y As Primitive)
            If imageName.IsEmpty Then
                Return
            End If
            VerifyAccess()
            Dim image1 As BitmapSource = ImageList.GetBitmap(imageName)
            If image1 IsNot Nothing Then
                BeginInvoke(
                    Sub()
                        Dim drawingContext1 As DrawingContext = _mainDrawing.Append()
                        drawingContext1.DrawImage(image1, New Windows.Rect(x, y, image1.PixelWidth, image1.PixelHeight))
                        drawingContext1.Close()
                        AddRasterizeOperationToQueue()
                    End Sub)
            End If
        End Sub

        ''' <summary>
        ''' Draws the pixel specified by the x and y co-ordinates using the specified color.
        ''' </summary>
        ''' <param name="x">
        ''' The x co-ordinate of the pixel.
        ''' </param>
        ''' <param name="y">
        ''' The y co-ordinate of the pixel.
        ''' </param>
        ''' <param name="color">
        ''' The color of the pixel to set.
        ''' </param>
        Public Shared Sub SetPixel(x As Primitive, y As Primitive, color1 As Primitive)
            VerifyAccess()
            BeginInvoke(
                Sub()
                    Dim drawingContext1 = _mainDrawing.Append()
                    drawingContext1.DrawRectangle(New SolidColorBrush(GetColorFromString(color1)), Nothing, New Windows.Rect(x, y, 1.0, 1.0))
                    drawingContext1.Close()
                    AddRasterizeOperationToQueue()
                End Sub)
        End Sub

        ''' <summary>
        ''' Gets the color of the pixel at the specified x and y co-ordinates.
        ''' </summary>
        ''' <param name="x">
        ''' The x co-ordinate of the pixel.
        ''' </param>
        ''' <param name="y">
        ''' The y co-ordinate of the pixel.
        ''' </param>
        ''' <returns>
        ''' The color of the pixel.
        ''' </returns>
        Public Shared Function GetPixel(x As Primitive, y As Primitive) As Primitive
            VerifyAccess()
            If x < 0 Then x = 0
            If y < 0 Then y = 0
            If x > Width Then x = Width
            If y > Height Then y = Height

            Return CType(InvokeWithReturn(
                Function() As Primitive
                    Rasterize()
                    Dim array As Byte() = New Byte(3) {}
                    Dim stride = CInt(_renderBitmap.Width * (_renderBitmap.Format.BitsPerPixel + 7) / 8)
                    _renderBitmap.CopyPixels(New Int32Rect(x, y, 1, 1), array, stride, 0)
                    Return CType($"#{array(2):X2}{array(1):X2}{array(0):X2}", Primitive)
                End Function), Primitive)
        End Function
        ''' <summary>
        ''' Gets a valid random color.
        ''' </summary>
        ''' <returns>
        ''' A valid random color.
        ''' </returns>
        Public Shared Function GetRandomColor() As Primitive
            If _random Is Nothing Then
                _random = New Random(Now.Ticks Mod Integer.MaxValue)
            End If

            Return $"#{_random.[Next](256):X2}{_random.[Next](256):X2}{_random.[Next](256):X2}"
        End Function
        ''' <summary>
        ''' Constructs a color given the Red, Green and Blue values.
        ''' </summary>
        ''' <param name="red">
        ''' The red component of the Color (0-255).
        ''' </param>
        ''' <param name="green">
        ''' The green component of the color (0-255).
        ''' </param>
        ''' <param name="blue">
        ''' The blue component of the color (0-255).
        ''' </param>
        ''' <returns>
        ''' Returns a color that can be used to set the brush or pen color.
        ''' </returns>
        Public Shared Function GetColorFromRGB(red As Primitive, green As Primitive, blue As Primitive) As Primitive
            Dim num As Integer = Math.Abs(CInt(red) Mod 256)
            Dim num2 As Integer = Math.Abs(CInt(green) Mod 256)
            Dim num3 As Integer = Math.Abs(CInt(blue) Mod 256)
            Return $"#{num:X2}{num2:X2}{num3:X2}"
        End Function

        ''' <summary>
        ''' Clears the window.
        ''' </summary>
        Public Shared Sub Clear()
            VerifyAccess()
            Invoke(Sub()
                       _mainCanvas.Children.Clear()
                       _renderBitmap.Clear()
                       _mainDrawing.Children.Clear()
                       _objectsMap.Clear()
                   End Sub)
        End Sub
        ''' <summary>
        ''' Displays a message box to the user.
        ''' </summary>
        ''' <param name="text">
        ''' The text to be displayed on the message box.
        ''' </param>
        ''' <param name="title">
        ''' The title for the message box.
        ''' </param>
        Public Shared Sub ShowMessage(text As Primitive, title1 As Primitive)
            VerifyAccess()
            Invoke(Sub()
                       MessageBox.Show(_window, text, title1)
                   End Sub)
        End Sub
        Private Shared Sub CreateWindow()
            SyncLock _syncLock
                Invoke(Sub()
                           _window = New Window
                           _windowCreated = True
                           _window.Title = "Small Basic Graphics Window"
                           If _backgroundBrush Is Nothing Then _backgroundBrush = Media.Brushes.White
                           _window.Background = _backgroundBrush
                           If _fillBrush Is Nothing Then _fillBrush = Media.Brushes.CornflowerBlue
                           If _pen Is Nothing Then _pen = New Media.Pen(System.Windows.Media.Brushes.Black, 2.0)
                           If _fontFamily Is Nothing Then _fontFamily = New Media.FontFamily("Tahoma")
                           _window.Height = 480.0
                           _window.Width = 640.0
                           If Not _isHidden Then _window.Show()

                           AddHandler _window.SourceInitialized,
                           Sub()
                               Dim handle1 As IntPtr = CType(PresentationSource.FromVisual(_window), HwndSource).Handle
                               Dim windowLong As UInteger = NativeHelper.GetWindowLong(handle1, -16)
                               windowLong = (windowLong And &HFFFEFFFFUI) Or &H20000UI
                               NativeHelper.SetWindowLong(handle1, -16, windowLong)
                               Dim rect2 As Internal.RECT = Nothing
                               rect2.Left = 0
                               rect2.Top = 0
                               rect2.Right = _window.Width
                               rect2.Bottom = _window.Height
                               Dim lpRect As Internal.RECT = rect2
                               NativeHelper.AdjustWindowRect(lpRect, windowLong, bMenu:=False)
                               _window.Width = lpRect.Right - lpRect.Left
                               _window.Height = lpRect.Bottom - lpRect.Top
                               NativeHelper.SetForegroundWindow(handle1)
                           End Sub
                           SetWindowContent()
                           SignupForWindowEvents()
                           _window.UpdateLayout()
                       End Sub)
            End SyncLock
        End Sub

        Private Shared Sub SignupForWindowEvents()
            AddHandler _window.SizeChanged, AddressOf WindowSizeChanged
            AddHandler _window.Closing, AddressOf WindowClosing
            AddHandler _window.MouseDown, Sub() RaiseEvent MouseDown()
            AddHandler _window.MouseUp, Sub() RaiseEvent MouseUp()

            AddHandler _window.MouseMove,
                Sub(sender As Object, e As MouseEventArgs)
                    Dim position As Windows.Point = e.GetPosition(_mainCanvas)
                    _mouseX = position.X
                    _mouseY = position.Y
                    RaiseEvent MouseMove()
                End Sub

            AddHandler _window.KeyDown,
                Sub(sender As Object, e As KeyEventArgs)
                    _lastKey = e.Key
                    BeginInvoke(Sub() RaiseEvent KeyDown())
                End Sub

            AddHandler _window.TextInput,
                Sub(sender As Object, e As TextCompositionEventArgs)
                    _lastText = e.Text
                    BeginInvoke(Sub() RaiseEvent TextInput())
                End Sub

            AddHandler _window.KeyUp,
                Sub(sender As Object, e As KeyEventArgs)
                    _lastKey = e.Key
                    BeginInvoke(Sub() RaiseEvent KeyUp())
                End Sub

        End Sub

        Private Shared Sub SetWindowContent()
            _mainCanvas = New Canvas
            _mainCanvas.HorizontalAlignment = HorizontalAlignment.Stretch
            _mainCanvas.VerticalAlignment = VerticalAlignment.Stretch
            _renderBitmap = New RenderTargetBitmap(_window.Width + 120, _window.Height + 120, 96.0, 96.0, PixelFormats.[Default])
            Dim image1 As New Windows.Controls.Image
            image1.Source = _renderBitmap
            image1.Stretch = Stretch.None
            _bitmapContainer = image1
            _mainDrawing = New DrawingGroup
            _visualContainer = New VisualContainer(_mainDrawing)
            Dim grid1 As New Grid
            grid1.Children.Add(_bitmapContainer)
            grid1.Children.Add(_visualContainer)
            grid1.Children.Add(_mainCanvas)
            _window.Content = grid1
        End Sub

        Private Shared Sub WindowClosing(sender As Object, e As CancelEventArgs)
            SmallBasicApplication.End()
        End Sub

        Private Shared Sub WindowSizeChanged(sender As Object, e As SizeChangedEventArgs)
            If e.NewSize.Width > _renderBitmap.Width OrElse e.NewSize.Height > _renderBitmap.Height Then
                Dim renderBitmap As RenderTargetBitmap = _renderBitmap
                _renderBitmap = New RenderTargetBitmap(e.NewSize.Width + 120, e.NewSize.Height + 120, 96.0, 96.0, PixelFormats.[Default])
                Dim drawingVisual1 As New DrawingVisual
                Dim drawingContext1 As DrawingContext = drawingVisual1.RenderOpen()
                drawingContext1.DrawImage(renderBitmap, New Windows.Rect(0.0, 0.0, renderBitmap.Width, renderBitmap.Height))
                drawingContext1.Close()
                _renderBitmap.Render(drawingVisual1)
                _bitmapContainer.Width = _renderBitmap.Width
                _bitmapContainer.Height = _renderBitmap.Height
                _bitmapContainer.Source = _renderBitmap
                _bitmapContainer.Stretch = Stretch.None
            End If
        End Sub
        Friend Shared Sub VerifyAccess()
            If Not _windowCreated Then
                CreateWindow()
                If Not _isHidden Then
                    Show()
                End If
            End If
        End Sub
        Friend Shared Sub BeginInvoke(invokeDelegate As InvokeHelper)
            SmallBasicApplication.BeginInvoke(invokeDelegate)
        End Sub
        Friend Shared Sub Invoke(invokeDelegate As InvokeHelper)
            SmallBasicApplication.Invoke(invokeDelegate)
        End Sub
        Friend Shared Function InvokeWithReturn(invokeDelegate As InvokeHelperWithReturn) As Object
            Return SmallBasicApplication.InvokeWithReturn(invokeDelegate)
        End Function
        Friend Shared Sub AddShape(name As String, shape As FrameworkElement)
            VerifyAccess()
            shape.Name = name
            _objectsMap(name) = shape
            _mainCanvas.Children.Add(shape)
        End Sub
        Friend Shared Sub AddControl(name As String, control1 As Control)
            VerifyAccess()
            control1.Foreground = _fillBrush
            control1.FontFamily = _fontFamily
            control1.FontStyle = _fontStyle
            control1.FontSize = _fontSize
            control1.FontWeight = _fontWeight
            AddShape(name, control1)
        End Sub
        Friend Shared Sub RemoveShape(name As Primitive)
            Dim shape As UIElement = Nothing
            If _objectsMap.TryGetValue(name, shape) Then
                _objectsMap.Remove(name)
                Invoke(Sub()
                           _mainCanvas.Children.Remove(shape)
                       End Sub)
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
        Friend Shared Sub DoubleAnimateProperty(animatable As IAnimatable, [property] As DependencyProperty, [end] As Double, aTimespan As Double)
            Dim num As Double = CDbl(CType(animatable, DependencyObject).GetValue([property]))
            If Double.IsNaN(num) Then num = 0.0
            Dim doubleAnimation1 As New DoubleAnimation(
                    num,
                    [end],
                    New Duration(TimeSpan.FromMilliseconds(aTimespan))
                )
            doubleAnimation1.FillBehavior = FillBehavior.HoldEnd
            doubleAnimation1.DecelerationRatio = 0.2
            Dim animation As DoubleAnimation = doubleAnimation1
            animatable.BeginAnimation([property], animation)
        End Sub

        Friend Shared Function GetColorFromString(color1 As String) As Windows.Media.Color
            Try
                Return CType(Windows.Media.ColorConverter.ConvertFromString(color1), Windows.Media.Color)
            Catch __unusedException1__ As Exception

            End Try

            Return Colors.Black
        End Function

        Friend Shared Function GetStringFromColor(color1 As Windows.Media.Color) As String
            Dim c As System.Drawing.Color = System.Drawing.Color.FromArgb(color1.A, color1.R, color1.G, color1.B)
            Return ColorTranslator.ToHtml(c)
        End Function

        Friend Shared Sub SetCursor(cursor1 As Cursor)
            VerifyAccess()
            Invoke(Sub()
                       _window.Cursor = cursor1
                   End Sub)
        End Sub
    End Class
End Namespace
