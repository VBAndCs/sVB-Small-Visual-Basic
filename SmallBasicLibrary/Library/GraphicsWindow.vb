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
        Friend Shared _fillBrush As SolidColorBrush = Media.Brushes.SlateBlue
        Friend Shared _fontFamily As Media.FontFamily
        Friend Shared _fontSize As Double = 12.0
        Friend Shared _fontStyle As System.Windows.FontStyle = FontStyles.Normal
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
        Private Shared _bitmapContainer As System.Windows.Controls.Image
        Private Shared _isHidden As Boolean
        Private Shared _syncLock As New Object


        ''' <summary>
        ''' Gets or sets the Background color of the Graphics Window.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Color)>
        Public Shared Property BackgroundColor As Primitive
            Get
                VerifyAccess()
                Return CStr(InvokeWithReturn(Function() GetStringFromColor(_backgroundBrush.Color)))
            End Get

            Set(Value As Primitive)
                VerifyAccess()
                BeginInvoke(
                    Sub()
                        _backgroundBrush = New SolidColorBrush(GetColorFromString(Value))
                        _window.Background = _backgroundBrush
                    End Sub)
            End Set
        End Property

        ''' <summary>
        ''' Gets or sets the brush color to be used to fill shapes drawn on the Graphics Window.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Color)>
        Public Shared Property BrushColor As Primitive
            Get
                VerifyAccess()
                If _fillBrush Is Nothing Then
                    Return WinForms.Colors.None
                End If

                Return CStr(InvokeWithReturn(
                    Function() GetStringFromColor(_fillBrush.Color))
                )
            End Get

            Set(Value As Primitive)
                VerifyAccess()
                BeginInvoke(
                    Sub()
                        If WinForms.Color.IsNone(Value) Then
                            _fillBrush = Nothing
                        Else
                            _fillBrush = New SolidColorBrush(GetColorFromString(Value))
                            _fillBrush.Freeze()
                        End If
                    End Sub)
            End Set
        End Property

        ''' <summary>
        ''' Specifies whether or not the Graphics Window can be resized by the user.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Boolean)>
        Public Shared Property CanResize As Primitive
            Get
                VerifyAccess()
                Return CBool(InvokeWithReturn(Function() _window.ResizeMode = ResizeMode.CanResize))
            End Get

            Set(Value As Primitive)
                VerifyAccess()
                BeginInvoke(
                    Sub()
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
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Property PenWidth As Primitive
            Get
                VerifyAccess()
                Return CDbl(InvokeWithReturn(
                        Function() If(_pen IsNot Nothing, _pen.Thickness, 2.0)
                ))
            End Get

            Set(Value As Primitive)
                VerifyAccess()
                BeginInvoke(Sub()
                                _pen = New Media.Pen(_pen.Brush, Value)
                                _pen.Freeze()
                            End Sub)
            End Set
        End Property

        ''' <summary>
        ''' Gets or sets the color of the pen used to draw shapes on the Graphics Window.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Color)>
        Public Shared Property PenColor As Primitive
            Get
                VerifyAccess()
                Return New Primitive(InvokeWithReturn(
                    Function() If(
                        _pen IsNot Nothing,
                        GetStringFromColor(CType(_pen.Brush, SolidColorBrush).Color),
                        WinForms.Colors.None.AsString())
                    ))
            End Get

            Set(Value As Primitive)
                VerifyAccess()
                BeginInvoke(
                    Sub()
                        If WinForms.Color.IsNone(Value) Then
                            _pen = Nothing
                        Else
                            Dim colorFromString As Media.Color = GetColorFromString(Value)
                            _pen = New Media.Pen(New SolidColorBrush(colorFromString), _pen.Thickness)
                            _pen.Freeze()
                        End If
                    End Sub)
            End Set
        End Property

        ''' <summary>
        ''' Gets or sets the Font Name to be used when drawing text on the Graphics Window.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Property FontName As Primitive
            Get
                Return CStr(InvokeWithReturn(Function() If((_fontFamily IsNot Nothing), _fontFamily.Source, "Tahoma")))
            End Get

            Set(Value As Primitive)
                BeginInvoke(Sub()
                                _fontFamily = New Media.FontFamily(Value)
                            End Sub)
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
                BeginInvoke(Sub()
                                _fontSize = Value
                            End Sub)
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
        <WinForms.ReturnValueType(VariableType.Boolean)>
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
        <WinForms.ReturnValueType(VariableType.String)>
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
        <WinForms.ReturnValueType(VariableType.Double)>
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
        <WinForms.ReturnValueType(VariableType.Double)>
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
        <WinForms.ReturnValueType(VariableType.Double)>
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
        <WinForms.ReturnValueType(VariableType.Double)>
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
        <WinForms.ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property LastKey As Primitive
            Get
                Return _lastKey.ToString()
            End Get
        End Property

        ''' <summary>
        ''' Gets the last text that was entered on the Graphics Window.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared ReadOnly Property LastText As Primitive
            Get
                Return _lastText.ToString()
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

        Private Shared Events As New EventHandlerList

        ''' <summary>
        ''' Raises an event when a key is pressed down on the keyboard.
        ''' </summary>
        Public Shared Custom Event KeyDown As SmallVisualBasicCallback
            AddHandler(Value As SmallVisualBasicCallback)
                VerifyAccess()
                Dim Key = NameOf(KeyDown)
                Dim h = TryCast(Events(Key), SmallVisualBasicCallback)
                If h IsNot Nothing Then Events.RemoveHandler(Key, h)
                Events.AddHandler(Key, Value)
            End AddHandler

            RemoveHandler(Value As SmallVisualBasicCallback)
                Events.RemoveHandler(NameOf(KeyDown), Value)
            End RemoveHandler

            RaiseEvent()
                Dim h = TryCast(Events(NameOf(KeyDown)), SmallVisualBasicCallback)
                If h IsNot Nothing Then h.Invoke()
            End RaiseEvent
        End Event

        ''' <summary>
        ''' Raises an event when a key is released on the keyboard.
        ''' </summary>
        Public Shared Custom Event KeyUp As SmallVisualBasicCallback
            AddHandler(Value As SmallVisualBasicCallback)
                VerifyAccess()
                Dim Key = NameOf(KeyUp)
                Dim h = TryCast(Events(Key), SmallVisualBasicCallback)
                If h IsNot Nothing Then Events.RemoveHandler(Key, h)
                Events.AddHandler(Key, Value)
            End AddHandler

            RemoveHandler(Value As SmallVisualBasicCallback)
                Events.RemoveHandler(NameOf(KeyUp), Value)
            End RemoveHandler

            RaiseEvent()
                Dim h = TryCast(Events(NameOf(KeyUp)), SmallVisualBasicCallback)
                If h IsNot Nothing Then h.Invoke()
            End RaiseEvent
        End Event

        ''' <summary>
        ''' Raises an event when the mouse button is clicked down.
        ''' </summary>
        Public Shared Custom Event MouseDown As SmallVisualBasicCallback
            AddHandler(Value As SmallVisualBasicCallback)
                VerifyAccess()
                Dim Key = NameOf(MouseDown)
                Dim h = TryCast(Events(Key), SmallVisualBasicCallback)
                If h IsNot Nothing Then Events.RemoveHandler(Key, h)
                Events.AddHandler(Key, Value)
            End AddHandler

            RemoveHandler(Value As SmallVisualBasicCallback)
                Events.RemoveHandler(NameOf(MouseDown), Value)
            End RemoveHandler

            RaiseEvent()
                Dim h = TryCast(Events(NameOf(MouseDown)), SmallVisualBasicCallback)
                If h IsNot Nothing Then h.Invoke()
            End RaiseEvent

        End Event

        ''' <summary>
        ''' Raises an event when the mouse button is released.
        ''' </summary>
        Public Shared Custom Event MouseUp As SmallVisualBasicCallback
            AddHandler(Value As SmallVisualBasicCallback)
                VerifyAccess()
                Dim Key = NameOf(MouseUp)
                Dim h = TryCast(Events(Key), SmallVisualBasicCallback)
                If h IsNot Nothing Then Events.RemoveHandler(Key, h)
                Events.AddHandler(Key, Value)
            End AddHandler

            RemoveHandler(Value As SmallVisualBasicCallback)
                Events.RemoveHandler(NameOf(MouseUp), Value)
            End RemoveHandler

            RaiseEvent()
                Dim h = TryCast(Events(NameOf(MouseUp)), SmallVisualBasicCallback)
                If h IsNot Nothing Then h.Invoke()
            End RaiseEvent

        End Event

        ''' <summary>
        ''' Raises an event when the mouse is moved around.
        ''' </summary>
        Public Shared Custom Event MouseMove As SmallVisualBasicCallback
            AddHandler(Value As SmallVisualBasicCallback)
                VerifyAccess()
                Dim Key = NameOf(MouseMove)
                Dim h = TryCast(Events(Key), SmallVisualBasicCallback)
                If h IsNot Nothing Then Events.RemoveHandler(Key, h)
                Events.AddHandler(Key, Value)
            End AddHandler

            RemoveHandler(Value As SmallVisualBasicCallback)
                Events.RemoveHandler(NameOf(MouseMove), Value)
            End RemoveHandler

            RaiseEvent()
                Dim h = TryCast(Events(NameOf(MouseMove)), SmallVisualBasicCallback)
                If h IsNot Nothing Then h.Invoke()
            End RaiseEvent


        End Event

        ''' <summary>
        ''' Raises an event when text is entered on the GraphicsWindow.
        ''' </summary>
        Public Shared Custom Event TextInput As SmallVisualBasicCallback
            AddHandler(Value As SmallVisualBasicCallback)
                VerifyAccess()
                Dim Key = NameOf(TextInput)
                Dim h = TryCast(Events(Key), SmallVisualBasicCallback)
                If h IsNot Nothing Then Events.RemoveHandler(Key, h)
                Events.AddHandler(Key, Value)
            End AddHandler

            RemoveHandler(Value As SmallVisualBasicCallback)
                Events.RemoveHandler(NameOf(TextInput), Value)
            End RemoveHandler

            RaiseEvent()
                Dim h = TryCast(Events(NameOf(TextInput)), SmallVisualBasicCallback)
                If h IsNot Nothing Then h.Invoke()
            End RaiseEvent


        End Event

        ''' <summary>
        ''' Shows the Graphics window to enable interactions with it.
        ''' </summary>
        Public Shared Sub Show()
            If Not _windowCreated Then CreateWindow()

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
            BeginInvoke(
                Sub()
                    Dim dc As DrawingContext = _mainDrawing.Append()
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
            BeginInvoke(
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
            BeginInvoke(
                Sub()
                    Dim drawingContext = _mainDrawing.Append()
                    Dim w = CDbl(width) / 2
                    Dim h = CDbl(height) / 2
                    drawingContext.DrawEllipse(
                        Nothing,
                        _pen, New System.Windows.Point(CDbl(x) + w, CDbl(y) + h),
                        w,
                        h
                    )
                    drawingContext.Close()
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
            BeginInvoke(Sub()
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
        Public Shared Sub DrawTriangle(x1 As Primitive, y1 As Primitive, x2 As Primitive, y2 As Primitive, x3 As Primitive, y3 As Primitive)
            VerifyAccess()
            BeginInvoke(Sub()
                            Dim dc As DrawingContext = _mainDrawing.Append()
                            Dim value As New PathFigure With
                            {
                            .StartPoint = New System.Windows.Point(x1, y1),
                            .IsClosed = True,
                            .Segments = New PathSegmentCollection From {
                                    CType(New LineSegment(New System.Windows.Point(x2, y2), isStroked:=True), PathSegment),
                                    CType(New LineSegment(New System.Windows.Point(x3, y3), isStroked:=True), PathSegment)
                            }
                            }
                            Dim figures As New PathFigureCollection From {
                                value
                            }
                            Dim pathGeometry1 As New PathGeometry
                            pathGeometry1.Figures = figures
                            pathGeometry1.Freeze()
                            dc.DrawGeometry(Nothing, _pen, pathGeometry1)
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
            BeginInvoke(Sub()
                            Dim dc As DrawingContext = _mainDrawing.Append()
                            Dim value As New PathFigure With
                            {
                            .StartPoint = New System.Windows.Point(x1, y1),
                            .IsClosed = True,
                            .Segments = New PathSegmentCollection From {
                                    CType(New LineSegment(New System.Windows.Point(x2, y2), isStroked:=True), PathSegment),
                                    CType(New LineSegment(New System.Windows.Point(x3, y3), isStroked:=True), PathSegment)
                            }
                            }
                            Dim figures As New PathFigureCollection From {
                                value
                            }
                            Dim pathGeometry1 As New PathGeometry
                            pathGeometry1.Figures = figures
                            pathGeometry1.Freeze()
                            dc.DrawGeometry(_fillBrush, Nothing, pathGeometry1)
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
            BeginInvoke(
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
                BeginInvoke(Sub()
                                Dim dc As DrawingContext = _mainDrawing.Append()
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
                BeginInvoke(
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
                        Dim dc As DrawingContext = _mainDrawing.Append()
                        dc.DrawImage(image1, New System.Windows.Rect(x, y, width1, height1))
                        dc.Close()
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
                        Dim dc As DrawingContext = _mainDrawing.Append()
                        dc.DrawImage(image1, New System.Windows.Rect(x, y, image1.PixelWidth, image1.PixelHeight))
                        dc.Close()
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
                    Dim dc = _mainDrawing.Append()
                    dc.DrawRectangle(New SolidColorBrush(GetColorFromString(color1)), Nothing, New System.Windows.Rect(x, y, 1.0, 1.0))
                    dc.Close()
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
        <WinForms.ReturnValueType(VariableType.Color)>
        Public Shared Function GetPixel(x As Primitive, y As Primitive) As Primitive
            VerifyAccess()
            If x < 0 Then x = 0
            If y < 0 Then y = 0
            If x > Width Then x = Width
            If y > Height Then y = Height

            Return CType(InvokeWithReturn(
                Function() As Primitive
                    Rasterize()
                    Dim colorBGR As Byte() = New Byte(3) {}
                    Dim stride = CInt(_renderBitmap.Width * (_renderBitmap.Format.BitsPerPixel + 7) / 8)
                    _renderBitmap.CopyPixels(New Int32Rect(x, y, 1, 1), colorBGR, stride, 0)
                    Return CType($"#{colorBGR(2):X2}{colorBGR(1):X2}{colorBGR(0):X2}", Primitive)
                End Function), Primitive)
        End Function

        ''' <summary>
        ''' Gets a valid random color.
        ''' </summary>
        ''' <returns>
        ''' A valid random color.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.Color)>
        Public Shared Function GetRandomColor() As Primitive
            If _random Is Nothing Then
                _random = New Random(Now.Ticks Mod Integer.MaxValue)
            End If

            Return $"#{_random.Next(256):X2}{_random.Next(256):X2}{_random.Next(256):X2}"
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
        <WinForms.ReturnValueType(VariableType.Color)>
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
        Public Shared Sub ShowMessage(text As Primitive, title As Primitive)
            VerifyAccess()
            Invoke(Sub()
                       MessageBox.Show(_window, text, title)
                   End Sub)
        End Sub

        Private Shared Sub CreateWindow()
            SyncLock _syncLock
                Invoke(
                    Sub()
                        _window = New Window()
                        _windowCreated = True
                        _window.Title = "Small Visual Basic Graphics Window"
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
                        WinForms.Forms._forms(Controls.GW_NAME) = _window
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
                    Dim position As System.Windows.Point = e.GetPosition(_mainCanvas)
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
            _renderBitmap = New RenderTargetBitmap(_window.Width + 120, _window.Height + 120, 96.0, 96.0, PixelFormats.Default)
            Dim image1 As New System.Windows.Controls.Image
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
            If WinForms.Forms._forms.Count = 0 AndAlso
                Not TextWindow._windowVisible Then
                SmallBasicApplication.End()
            End If

            WinForms.Forms.RemoveFormAndControls(Controls.GW_NAME)
        End Sub

        Private Shared Sub WindowSizeChanged(sender As Object, e As SizeChangedEventArgs)
            If e.NewSize.Width > _renderBitmap.Width OrElse e.NewSize.Height > _renderBitmap.Height Then
                Dim renderBitmap As RenderTargetBitmap = _renderBitmap
                _renderBitmap = New RenderTargetBitmap(e.NewSize.Width + 120, e.NewSize.Height + 120, 96.0, 96.0, PixelFormats.Default)
                Dim drawingVisual1 As New DrawingVisual
                Dim dc As DrawingContext = drawingVisual1.RenderOpen()
                dc.DrawImage(renderBitmap, New System.Windows.Rect(0.0, 0.0, renderBitmap.Width, renderBitmap.Height))
                dc.Close()
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

        Friend Shared Sub AddShape(
                           name As String,
                           shape As FrameworkElement,
                           Optional addToCanvas As Boolean = True
                    )

            VerifyAccess()
            If name.StartsWith(Controls.GW_NAME) Then
                shape.Name = name.Substring(Controls.GW_NAME.Length + 1)
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

        Friend Shared Sub DoubleAnimateProperty(obj As IAnimatable, [property] As DependencyProperty, [end] As Double, duration As Double)
            Dim dpo = CType(obj, DependencyObject)
            Dim start = CDbl(dpo.GetValue([property]))
            If Double.IsNaN(start) Then start = 0.0

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
            Catch __unusedException1__ As Exception

            End Try

            Return Colors.Black
        End Function

        Friend Shared Function GetStringFromColor(color1 As Media.Color) As String
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
