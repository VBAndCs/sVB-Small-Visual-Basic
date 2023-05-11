Imports System.Text
Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.Windows.Controls.Primitives

<TemplatePart(Name:="PART_CurrentColor", Type:=GetType(TextBox))> _
Friend Class ColorPicker
    Inherits Control

    Friend Const PART_CurrentColor As String = "PART_CurrentColor"
    Friend _HSBSetInternally As Boolean = False
    Friend _RGBSetInternally As Boolean = False
    Friend _BrushSetInternally As Boolean = False
    Friend _BrushTypeSetInternally As Boolean = False
    Friend _UpdateBrush As Boolean = True
    Private privateCurrentColorTextBox As TextBox

    Friend Property CurrentColorTextBox() As TextBox
        Get
            Return privateCurrentColorTextBox
        End Get
        Private Set(ByVal value As TextBox)
            privateCurrentColorTextBox = value
        End Set
    End Property

    Shared Sub New()
        DefaultStyleKeyProperty.OverrideMetadata(GetType(ColorPicker), New FrameworkPropertyMetadata(GetType(ColorPicker)))
    End Sub

    Public Shared RemoveGradientStop As New RoutedCommand()
    Public Shared ReverseGradientStop As New RoutedCommand()

    Public Property ShowPopUpColorDegrees As Boolean
        Get
            Return GetValue(ShowPopUpColorDegreesProperty)
        End Get

        Set(ByVal value As Boolean)
            SetValue(ShowPopUpColorDegreesProperty, value)
        End Set
    End Property

    Public Shared ReadOnly ShowPopUpColorDegreesProperty As DependencyProperty = _
                           DependencyProperty.Register("ShowPopUpColorDegrees", _
                           GetType(Boolean), GetType(ColorPicker), _
                           New PropertyMetadata(True))

    Dim Scv As ScrollViewer
    Dim TransPker As TransformPicker
    Dim ColorLst As ColorList

    Public Overrides Sub OnApplyTemplate()
        MyBase.OnApplyTemplate()
        ColorLst = GetTemplateChild("PART_ColorList")

        Scv = GetTemplateChild("PART_ScrollViewer")
        AddHandler Scv.ScrollChanged, AddressOf Scv_ScrollChanged

        TransPker = GetTemplateChild("PART_TransformPicker")
        AddHandler TransPker.TransformChanged, AddressOf TransPker_TransformChanged

        CurrentColorTextBox = TryCast(GetTemplateChild(PART_CurrentColor), TextBox)
        If CurrentColorTextBox IsNot Nothing Then
            AddHandler CurrentColorTextBox.PreviewKeyDown, AddressOf CurrentColorTextBox_PreviewKeyDown
        End If

        Me.CommandBindings.Add(New CommandBinding(ColorPicker.RemoveGradientStop, AddressOf RemoveGradientStop_Executed))
        Me.CommandBindings.Add(New CommandBinding(ColorPicker.ReverseGradientStop, AddressOf ReverseGradientStop_Executed))
    End Sub

    Private Sub CurrentColorTextBox_PreviewKeyDown(ByVal sender As Object, ByVal e As KeyEventArgs)
        If e.Key = Key.Enter Then
            Dim be As BindingExpression = CurrentColorTextBox.GetBindingExpression(TextBox.TextProperty)
            If be IsNot Nothing Then be.UpdateSource()
            e.Handled = True
        End If
    End Sub

    Private Sub RemoveGradientStop_Executed(ByVal sender As Object, ByVal e As ExecutedRoutedEventArgs)
        If Me.Gradients IsNot Nothing AndAlso Me.Gradients.Count > 2 Then
            Me.Gradients.Remove(Me.SelectedGradient)
            Me.SetBrush()
        End If
    End Sub

    Private Sub ReverseGradientStop_Executed(ByVal sender As Object, ByVal e As ExecutedRoutedEventArgs)
        Me._UpdateBrush = False
        Me._BrushSetInternally = True
        Try
            For Each gs As GradientStop In Gradients
                gs.Offset = 1.0 - gs.Offset
            Next gs
        Catch
        End Try

        Me._UpdateBrush = True
        Me._BrushSetInternally = False
        Me.SetBrush()
    End Sub

    Private Sub InitTransform()
        If Me.Brush.Transform Is Nothing OrElse Me.Brush.Transform.Value.IsIdentity Then
            Me._BrushSetInternally = True

            Dim _tg As New TransformGroup()
            _tg.Children.Add(New RotateTransform())
            _tg.Children.Add(New ScaleTransform())
            _tg.Children.Add(New SkewTransform())
            _tg.Children.Add(New TranslateTransform())
            Me.Brush.Transform = _tg

            Me._BrushSetInternally = False
        End If
    End Sub

#Region "Private Properties"

    Private Property StartX() As Double
        Get
            Return CDbl(GetValue(StartXProperty))
        End Get
        Set(ByVal value As Double)
            SetValue(StartXProperty, value)
        End Set
    End Property

    Private Shared ReadOnly StartXProperty As DependencyProperty = DependencyProperty.Register("StartX", GetType(Double), GetType(ColorPicker), New PropertyMetadata(0.5, New PropertyChangedCallback(AddressOf StartXChanged)))
    Private Shared Sub StartXChanged(ByVal [property] As DependencyObject, ByVal args As DependencyPropertyChangedEventArgs)
        Dim cp As ColorPicker = TryCast([property], ColorPicker)
        If TypeOf cp.Brush Is LinearGradientBrush Then
            cp._BrushSetInternally = True
            TryCast(cp.Brush, LinearGradientBrush).StartPoint = New Point(CDbl(args.NewValue), (TryCast(cp.Brush, LinearGradientBrush)).StartPoint.Y)
            cp._BrushSetInternally = False
        End If
    End Sub

    Private Property StartY() As Double
        Get
            Return CDbl(GetValue(StartYProperty))
        End Get
        Set(ByVal value As Double)
            SetValue(StartYProperty, value)
        End Set
    End Property

    Private Shared ReadOnly StartYProperty As DependencyProperty = DependencyProperty.Register("StartY", GetType(Double), GetType(ColorPicker), New PropertyMetadata(0.0, New PropertyChangedCallback(AddressOf StartYChanged)))
    Private Shared Sub StartYChanged(ByVal [property] As DependencyObject, ByVal args As DependencyPropertyChangedEventArgs)
        Dim cp As ColorPicker = TryCast([property], ColorPicker)
        If TypeOf cp.Brush Is LinearGradientBrush Then
            cp._BrushSetInternally = True
            TryCast(cp.Brush, LinearGradientBrush).StartPoint = New Point((TryCast(cp.Brush, LinearGradientBrush)).StartPoint.X, CDbl(args.NewValue))
            cp._BrushSetInternally = False
        End If
    End Sub

    Private Property EndX() As Double
        Get
            Return CDbl(GetValue(EndXProperty))
        End Get
        Set(ByVal value As Double)
            SetValue(EndXProperty, value)
        End Set
    End Property

    Private Shared ReadOnly EndXProperty As DependencyProperty = DependencyProperty.Register("EndX", GetType(Double), GetType(ColorPicker), New PropertyMetadata(0.5, New PropertyChangedCallback(AddressOf EndXChanged)))
    Private Shared Sub EndXChanged(ByVal [property] As DependencyObject, ByVal args As DependencyPropertyChangedEventArgs)
        Dim cp As ColorPicker = TryCast([property], ColorPicker)
        If TypeOf cp.Brush Is LinearGradientBrush Then
            cp._BrushSetInternally = True
            TryCast(cp.Brush, LinearGradientBrush).EndPoint = New Point(CDbl(args.NewValue), (TryCast(cp.Brush, LinearGradientBrush)).EndPoint.Y)
            cp._BrushSetInternally = False
        End If
    End Sub

    Private Property EndY() As Double
        Get
            Return CDbl(GetValue(EndYProperty))
        End Get
        Set(ByVal value As Double)
            SetValue(EndYProperty, value)
        End Set
    End Property

    Private Shared ReadOnly EndYProperty As DependencyProperty = DependencyProperty.Register("EndY", GetType(Double), GetType(ColorPicker), New PropertyMetadata(1.0, New PropertyChangedCallback(AddressOf EndYChanged)))
    Private Shared Sub EndYChanged(ByVal [property] As DependencyObject, ByVal args As DependencyPropertyChangedEventArgs)
        Dim cp As ColorPicker = TryCast([property], ColorPicker)
        If TypeOf cp.Brush Is LinearGradientBrush Then
            cp._BrushSetInternally = True
            TryCast(cp.Brush, LinearGradientBrush).EndPoint = New Point((TryCast(cp.Brush, LinearGradientBrush)).EndPoint.X, CDbl(args.NewValue))
            cp._BrushSetInternally = False
        End If
    End Sub


    Private Property GradientOriginX() As Double
        Get
            Return CDbl(GetValue(GradientOriginXProperty))
        End Get
        Set(ByVal value As Double)
            SetValue(GradientOriginXProperty, value)
        End Set
    End Property

    Private Shared ReadOnly GradientOriginXProperty As DependencyProperty = DependencyProperty.Register("GradientOriginX", GetType(Double), GetType(ColorPicker), New PropertyMetadata(0.5, New PropertyChangedCallback(AddressOf GradientOriginXChanged)))

    Private Shared Sub GradientOriginXChanged(ByVal [property] As DependencyObject, ByVal args As DependencyPropertyChangedEventArgs)
        Dim cp As ColorPicker = TryCast([property], ColorPicker)
        If TypeOf cp.Brush Is RadialGradientBrush Then
            cp._BrushSetInternally = True
            TryCast(cp.Brush, RadialGradientBrush).GradientOrigin = New Point(CDbl(args.NewValue), (TryCast(cp.Brush, RadialGradientBrush)).GradientOrigin.Y)
            cp._BrushSetInternally = False
        End If
    End Sub

    Private Property GradientOriginY() As Double
        Get
            Return CDbl(GetValue(GradientOriginYProperty))
        End Get
        Set(ByVal value As Double)
            SetValue(GradientOriginYProperty, value)
        End Set
    End Property

    Private Shared ReadOnly GradientOriginYProperty As DependencyProperty = DependencyProperty.Register("GradientOriginY", GetType(Double), GetType(ColorPicker), New PropertyMetadata(0.5, New PropertyChangedCallback(AddressOf GradientOriginYChanged)))
    Private Shared Sub GradientOriginYChanged(ByVal [property] As DependencyObject, ByVal args As DependencyPropertyChangedEventArgs)
        Dim cp As ColorPicker = TryCast([property], ColorPicker)
        If TypeOf cp.Brush Is RadialGradientBrush Then
            cp._BrushSetInternally = True
            TryCast(cp.Brush, RadialGradientBrush).GradientOrigin = New Point((TryCast(cp.Brush, RadialGradientBrush)).GradientOrigin.X, CDbl(args.NewValue))
            cp._BrushSetInternally = False
        End If
    End Sub

    Private Property CenterX() As Double
        Get
            Return CDbl(GetValue(CenterXProperty))
        End Get
        Set(ByVal value As Double)
            SetValue(CenterXProperty, value)
        End Set
    End Property

    Private Shared ReadOnly CenterXProperty As DependencyProperty = DependencyProperty.Register("CenterX", GetType(Double), GetType(ColorPicker), New PropertyMetadata(0.5, New PropertyChangedCallback(AddressOf CenterXChanged)))
    Private Shared Sub CenterXChanged(ByVal [property] As DependencyObject, ByVal args As DependencyPropertyChangedEventArgs)
        Dim cp As ColorPicker = TryCast([property], ColorPicker)
        If TypeOf cp.Brush Is RadialGradientBrush Then
            cp._BrushSetInternally = True
            TryCast(cp.Brush, RadialGradientBrush).Center = New Point(CDbl(args.NewValue), (TryCast(cp.Brush, RadialGradientBrush)).Center.Y)
            cp._BrushSetInternally = False
        End If
    End Sub

    Private Property CenterY() As Double
        Get
            Return CDbl(GetValue(CenterYProperty))
        End Get
        Set(ByVal value As Double)
            SetValue(CenterYProperty, value)
        End Set
    End Property

    Private Shared ReadOnly CenterYProperty As DependencyProperty = DependencyProperty.Register("CenterY", GetType(Double), GetType(ColorPicker), New PropertyMetadata(0.5, New PropertyChangedCallback(AddressOf CenterYChanged)))
    Private Shared Sub CenterYChanged(ByVal [property] As DependencyObject, ByVal args As DependencyPropertyChangedEventArgs)
        Dim cp As ColorPicker = TryCast([property], ColorPicker)
        If TypeOf cp.Brush Is RadialGradientBrush Then
            cp._BrushSetInternally = True
            TryCast(cp.Brush, RadialGradientBrush).Center = New Point((TryCast(cp.Brush, RadialGradientBrush)).Center.X, CDbl(args.NewValue))
            cp._BrushSetInternally = False
        End If
    End Sub

    Private Property RadiusX() As Double
        Get
            Return CDbl(GetValue(RadiusXProperty))
        End Get
        Set(ByVal value As Double)
            SetValue(RadiusXProperty, value)
        End Set
    End Property

    Private Shared ReadOnly RadiusXProperty As DependencyProperty = DependencyProperty.Register("RadiusX", GetType(Double), GetType(ColorPicker), New PropertyMetadata(0.5, New PropertyChangedCallback(AddressOf RadiusXChanged)))

    Private Shared Sub RadiusXChanged(ByVal [property] As DependencyObject, ByVal args As DependencyPropertyChangedEventArgs)
        Dim cp As ColorPicker = TryCast([property], ColorPicker)
        If TypeOf cp.Brush Is RadialGradientBrush Then
            cp._BrushSetInternally = True
            TryCast(cp.Brush, RadialGradientBrush).RadiusX = CDbl(args.NewValue)
            cp._BrushSetInternally = False
        End If
    End Sub

    Private Property RadiusY() As Double
        Get
            Return CDbl(GetValue(RadiusYProperty))
        End Get
        Set(ByVal value As Double)
            SetValue(RadiusYProperty, value)
        End Set
    End Property

    Private Shared ReadOnly RadiusYProperty As DependencyProperty = DependencyProperty.Register("RadiusY", GetType(Double), GetType(ColorPicker), New PropertyMetadata(0.5, New PropertyChangedCallback(AddressOf RadiusYChanged)))
    Private Shared Sub RadiusYChanged(ByVal [property] As DependencyObject, ByVal args As DependencyPropertyChangedEventArgs)
        Dim cp As ColorPicker = TryCast([property], ColorPicker)
        If TypeOf cp.Brush Is RadialGradientBrush Then
            cp._BrushSetInternally = True
            TryCast(cp.Brush, RadialGradientBrush).RadiusY = CDbl(args.NewValue)
            cp._BrushSetInternally = False
        End If
    End Sub

    Private Property BrushOpacity() As Double
        Get
            Return CDbl(GetValue(BrushOpacityProperty))
        End Get
        Set(ByVal value As Double)
            SetValue(BrushOpacityProperty, value)
        End Set
    End Property

    Private Shared ReadOnly BrushOpacityProperty As DependencyProperty = DependencyProperty.Register("BrushOpacity", GetType(Double), GetType(ColorPicker), New PropertyMetadata(1.0))

    Private Property SpreadMethod() As GradientSpreadMethod
        Get
            Return CType(GetValue(SpreadMethodProperty), GradientSpreadMethod)
        End Get
        Set(ByVal value As GradientSpreadMethod)
            SetValue(SpreadMethodProperty, value)
        End Set
    End Property

    Private Shared ReadOnly SpreadMethodProperty As DependencyProperty = DependencyProperty.Register("SpreadMethod", GetType(GradientSpreadMethod), GetType(ColorPicker), New PropertyMetadata(GradientSpreadMethod.Pad, New PropertyChangedCallback(AddressOf SpreadMethodChanged)))
    Private Shared Sub SpreadMethodChanged(ByVal [property] As DependencyObject, ByVal args As DependencyPropertyChangedEventArgs)
        Dim cp As ColorPicker = TryCast([property], ColorPicker)
        If TypeOf cp.Brush Is GradientBrush Then
            cp._BrushSetInternally = True
            Try

                TryCast(cp.Brush, GradientBrush).SpreadMethod = CType(args.NewValue, GradientSpreadMethod)

            Catch ex As Exception

            End Try
            cp._BrushSetInternally = False
        End If
    End Sub

    Private Property ColorInterpolationMode() As ColorInterpolationMode
        Get
            Return CType(GetValue(ColorInterpolationModeProperty), ColorInterpolationMode)
        End Get
        Set(ByVal value As ColorInterpolationMode)
            SetValue(ColorInterpolationModeProperty, value)
        End Set
    End Property

    Private Shared ReadOnly ColorInterpolationModeProperty As DependencyProperty = DependencyProperty.Register("ColorInterpolationMode", GetType(ColorInterpolationMode), GetType(ColorPicker), New PropertyMetadata(ColorInterpolationMode.SRgbLinearInterpolation, New PropertyChangedCallback(AddressOf ColorInterpolationModeChanged)))

    Private Shared Sub ColorInterpolationModeChanged(ByVal [property] As DependencyObject, ByVal args As DependencyPropertyChangedEventArgs)
        Dim cp As ColorPicker = TryCast([property], ColorPicker)
        If TypeOf cp.Brush Is GradientBrush Then
            cp._BrushSetInternally = True
            TryCast(cp.Brush, GradientBrush).ColorInterpolationMode = CType(args.NewValue, ColorInterpolationMode)
            cp._BrushSetInternally = False
        End If
    End Sub

#End Region

#Region "Internal Properties"

    Friend Property Gradients() As ObservableCollection(Of GradientStop)
        Get
            Return CType(GetValue(GradientsProperty), ObservableCollection(Of GradientStop))
        End Get
        Set(ByVal value As ObservableCollection(Of GradientStop))
            SetValue(GradientsProperty, value)
        End Set
    End Property

    Friend Shared ReadOnly GradientsProperty As DependencyProperty = DependencyProperty.Register("Gradients", GetType(ObservableCollection(Of GradientStop)), GetType(ColorPicker))

    Friend Property SelectedGradient() As GradientStop
        Get
            Return CType(GetValue(SelectedGradientProperty), GradientStop)
        End Get
        Set(ByVal value As GradientStop)
            SetValue(SelectedGradientProperty, value)
        End Set
    End Property

    Friend Shared ReadOnly SelectedGradientProperty As DependencyProperty = DependencyProperty.Register("SelectedGradient", GetType(GradientStop), GetType(ColorPicker))

    Friend Property BrushType() As BrushTypes
        Get
            Return CType(GetValue(BrushTypeProperty), BrushTypes)
        End Get
        Set(ByVal value As BrushTypes)
            SetValue(BrushTypeProperty, value)
        End Set
    End Property

    Friend Shared ReadOnly BrushTypeProperty As DependencyProperty =
               DependencyProperty.Register(
                   "BrushType",
                   GetType(BrushTypes),
                   GetType(ColorPicker),
                   New FrameworkPropertyMetadata(
                       BrushTypes.None,
                       New PropertyChangedCallback(AddressOf BrushTypeChanged)
                   )
               )

    Private Shared Sub BrushTypeChanged(ByVal [property] As DependencyObject, ByVal args As DependencyPropertyChangedEventArgs)
        Dim c As ColorPicker = TryCast([property], ColorPicker)
        If Not c._BrushTypeSetInternally Then
            If c.Gradients Is Nothing Then
                c.Gradients = New ObservableCollection(Of GradientStop)()
                c.Gradients.Add(New GradientStop(c.Color, 0))
                c.Gradients.Add(New GradientStop(Colors.White, 1))
            End If

            c.SelectedGradient = c.Gradients(0)
            c.SetBrush(True)
        End If
    End Sub

#End Region

#Region "Public Properties"

    Public ReadOnly Iterator Property SpreadMethodTypes() As IEnumerable(Of System.Enum)
        Get
            Dim temp As GradientSpreadMethod = GradientSpreadMethod.Pad Or GradientSpreadMethod.Reflect Or GradientSpreadMethod.Repeat
            For Each value As [Enum] In [Enum].GetValues(temp.GetType())
                If temp.HasFlag(value) Then
                    Yield value
                End If
            Next value
        End Get
    End Property

    Public ReadOnly Iterator Property ColorInterpolationModeTypes() As IEnumerable(Of System.Enum)
        Get
            Dim temp As ColorInterpolationMode = ColorInterpolationMode.ScRgbLinearInterpolation Or ColorInterpolationMode.SRgbLinearInterpolation
            For Each value As [Enum] In [Enum].GetValues(temp.GetType())
                If temp.HasFlag(value) Then
                    Yield value
                End If
            Next value
        End Get
    End Property

    Public ReadOnly Iterator Property AvailableBrushTypes() As IEnumerable(Of System.Enum)
        Get
            Dim temp As BrushTypes = BrushTypes.None Or BrushTypes.Solid Or BrushTypes.Linear Or BrushTypes.Radial
            For Each value As [Enum] In [Enum].GetValues(temp.GetType())
                If temp.HasFlag(value) Then
                    Yield value
                End If
            Next value
        End Get
    End Property

    Public Property Brush() As Brush
        Get
            Return CType(GetValue(BrushProperty), Brush)
        End Get

        Set(ByVal value As Brush)
            Try
                SetValue(BrushProperty, value)
            Catch ex As Exception
            End Try
        End Set
    End Property

    Public Shared ReadOnly BrushProperty As DependencyProperty = DependencyProperty.Register("Brush", GetType(Brush), GetType(ColorPicker), New FrameworkPropertyMetadata(Nothing, New PropertyChangedCallback(AddressOf BrushChanged)))

    Private Shared Sub BrushChanged(ByVal [property] As DependencyObject, ByVal args As DependencyPropertyChangedEventArgs)
        Dim c As ColorPicker = TryCast([property], ColorPicker)
        Dim _brush As Brush = TryCast(args.NewValue, Brush)

        c._UpdateBrush = False
        If Not c._BrushSetInternally Then
            c._BrushTypeSetInternally = True

            If _brush Is Nothing Then
                c.BrushType = BrushTypes.None
            ElseIf TypeOf _brush Is SolidColorBrush Then
                c.BrushType = BrushTypes.Solid
                c.Color = TryCast(_brush, SolidColorBrush).Color
            ElseIf TypeOf _brush Is LinearGradientBrush Then
                Dim lgb As LinearGradientBrush = TryCast(_brush, LinearGradientBrush)
                'c.Opacity = lgb.Opacity;
                c.StartX = lgb.StartPoint.X
                c.StartY = lgb.StartPoint.Y
                c.EndX = lgb.EndPoint.X
                c.EndY = lgb.EndPoint.Y
                c.ColorInterpolationMode = lgb.ColorInterpolationMode
                c.SpreadMethod = lgb.SpreadMethod
                c.Gradients = New ObservableCollection(Of GradientStop)(lgb.GradientStops)
                c.SelectedGradient = c.Gradients(0)
                c.BrushType = BrushTypes.Linear
                c.Color = c.Gradients(0).Color
                If c.TransPker IsNot Nothing Then c.TransPker.Transform = _brush.RelativeTransform
            Else
                Dim rgb As RadialGradientBrush = TryCast(_brush, RadialGradientBrush)
                c.GradientOriginX = rgb.GradientOrigin.X
                c.GradientOriginY = rgb.GradientOrigin.Y
                c.RadiusX = rgb.RadiusX
                c.RadiusY = rgb.RadiusY
                c.CenterX = rgb.Center.X
                c.CenterY = rgb.Center.Y
                c.ColorInterpolationMode = rgb.ColorInterpolationMode
                c.SpreadMethod = rgb.SpreadMethod
                c.Gradients = New ObservableCollection(Of GradientStop)(rgb.GradientStops)
                c.BrushType = BrushTypes.Radial
                c.SelectedGradient = c.Gradients(0)
                c.Color = c.Gradients(0).Color
                If c.TransPker IsNot Nothing Then c.TransPker.Transform = _brush.RelativeTransform
            End If

            c._BrushTypeSetInternally = False
            c.RaiseColorChangedEvent(c.Color)
        End If
        c._UpdateBrush = True

    End Sub

    Public Property Color() As Color
        Get
            Return CType(GetValue(ColorProperty), Color)
        End Get
        Set(ByVal value As Color)
            SetValue(ColorProperty, value)
        End Set
    End Property

    Public Shared ReadOnly ColorProperty As DependencyProperty = DependencyProperty.Register("Color", GetType(Color), GetType(ColorPicker), New UIPropertyMetadata(Colors.Black, AddressOf OnColorChanged))

    Dim DontChangeColor As Boolean = False
    Public Shared Sub OnColorChanged(ByVal o As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
        Dim c As ColorPicker = CType(o, ColorPicker)
        If c.DontChangeColor Then Exit Sub

        If TypeOf e.NewValue Is Color Then
            c.DontChangeColor = True
            Dim _color As Color = CType(e.NewValue, Color)

            If Not c._HSBSetInternally Then
                ' update HSB value based on new value of color
                Dim H As Double = 0
                Dim S As Double = 0
                Dim B As Double = 0

                ColorHelper.HSBFromColor(_color, H, S, B)
                c._HSBSetInternally = True
                c.Alpha = CDbl(_color.A / 255.0R)
                c.Hue = H
                c.Saturation = S
                c.Brightness = B
                c._HSBSetInternally = False
                UpdateColorHSB(c, Nothing)
            End If

            If Not c._RGBSetInternally Then
                ' update RGB value based on new value of color
                c._RGBSetInternally = True
                c.A = _color.A
                c.R = _color.R
                c.G = _color.G
                c.B = _color.B
                c._RGBSetInternally = False
                UpdateColorRGB(c, Nothing)
            End If

            c.RaiseColorChangedEvent(CType(e.NewValue, Color))
            c.DontChangeColor = False
        End If
    End Sub

#End Region


#Region "Color Specific Properties"

    Private Property Hue() As Double
        Get
            Return CDbl(GetValue(HueProperty))
        End Get
        Set(ByVal value As Double)
            SetValue(HueProperty, value)
        End Set
    End Property
    Private Shared ReadOnly HueProperty As DependencyProperty = DependencyProperty.Register("Hue", GetType(Double), GetType(ColorPicker), New FrameworkPropertyMetadata(1.0, New PropertyChangedCallback(AddressOf UpdateColorHSB), New CoerceValueCallback(AddressOf HueCoerce)))
    Private Shared Function HueCoerce(ByVal d As DependencyObject, ByVal Hue As Object) As Object
        Dim v As Double = CDbl(Hue)
        If v < 0 Then
            Return 0.0
        End If
        If v > 1 Then
            Return 1.0
        End If
        Return v
    End Function


    Private Property Brightness() As Double
        Get
            Return CDbl(GetValue(BrightnessProperty))
        End Get
        Set(ByVal value As Double)
            SetValue(BrightnessProperty, value)
        End Set
    End Property
    Private Shared ReadOnly BrightnessProperty As DependencyProperty = DependencyProperty.Register("Brightness", GetType(Double), GetType(ColorPicker), New FrameworkPropertyMetadata(0.0, New PropertyChangedCallback(AddressOf UpdateColorHSB), New CoerceValueCallback(AddressOf BrightnessCoerce)))
    Private Shared Function BrightnessCoerce(ByVal d As DependencyObject, ByVal Brightness As Object) As Object
        Dim v As Double = CDbl(Brightness)
        If v < 0 Then
            Return 0.0
        End If
        If v > 1 Then
            Return 1.0
        End If
        Return v
    End Function


    Private Property Saturation() As Double
        Get
            Return CDbl(GetValue(SaturationProperty))
        End Get
        Set(ByVal value As Double)
            SetValue(SaturationProperty, value)
        End Set
    End Property
    Private Shared ReadOnly SaturationProperty As DependencyProperty = DependencyProperty.Register("Saturation", GetType(Double), GetType(ColorPicker), New FrameworkPropertyMetadata(0.0, New PropertyChangedCallback(AddressOf UpdateColorHSB), New CoerceValueCallback(AddressOf SaturationCoerce)))
    Private Shared Function SaturationCoerce(ByVal d As DependencyObject, ByVal Saturation As Object) As Object
        Dim v As Double = CDbl(Saturation)
        If v < 0 Then
            Return 0.0
        End If
        If v > 1 Then
            Return 1.0
        End If
        Return v
    End Function


    Private Property Alpha() As Double
        Get
            Return CDbl(GetValue(AlphaProperty))
        End Get
        Set(ByVal value As Double)
            SetValue(AlphaProperty, value)
        End Set
    End Property
    Private Shared ReadOnly AlphaProperty As DependencyProperty = DependencyProperty.Register("Alpha", GetType(Double), GetType(ColorPicker), New FrameworkPropertyMetadata(1.0, New PropertyChangedCallback(AddressOf UpdateColorHSB), New CoerceValueCallback(AddressOf AlphaCoerce)))
    Private Shared Function AlphaCoerce(ByVal d As DependencyObject, ByVal Alpha As Object) As Object
        Dim v As Double = CDbl(Alpha)
        If v < 0 Then
            Return 0.0
        End If
        If v > 1 Then
            Return 1.0
        End If
        Return v
    End Function


    Private Property A() As Integer
        Get
            Return CInt(Fix(GetValue(AProperty)))
        End Get
        Set(ByVal value As Integer)
            SetValue(AProperty, value)
        End Set
    End Property
    Private Shared ReadOnly AProperty As DependencyProperty = DependencyProperty.Register("A", GetType(Integer), GetType(ColorPicker), New FrameworkPropertyMetadata(255, New PropertyChangedCallback(AddressOf UpdateColorRGB), New CoerceValueCallback(AddressOf RGBCoerce)))


    Private Property R() As Integer
        Get
            Return CInt(Fix(GetValue(RProperty)))
        End Get
        Set(ByVal value As Integer)
            SetValue(RProperty, value)
        End Set
    End Property
    Private Shared ReadOnly RProperty As DependencyProperty = DependencyProperty.Register("R", GetType(Integer), GetType(ColorPicker), New FrameworkPropertyMetadata(New PropertyChangedCallback(AddressOf UpdateColorRGB), New CoerceValueCallback(AddressOf RGBCoerce)))


    Private Property G() As Integer
        Get
            Return CInt(Fix(GetValue(GProperty)))
        End Get
        Set(ByVal value As Integer)
            SetValue(GProperty, value)
        End Set
    End Property
    Private Shared ReadOnly GProperty As DependencyProperty = DependencyProperty.Register("G", GetType(Integer), GetType(ColorPicker), New FrameworkPropertyMetadata(New PropertyChangedCallback(AddressOf UpdateColorRGB), New CoerceValueCallback(AddressOf RGBCoerce)))


    Private Property B() As Integer
        Get
            Return CInt(Fix(GetValue(BProperty)))
        End Get
        Set(ByVal value As Integer)
            SetValue(BProperty, value)
        End Set
    End Property
    Private Shared ReadOnly BProperty As DependencyProperty = DependencyProperty.Register("B", GetType(Integer), GetType(ColorPicker), New FrameworkPropertyMetadata(New PropertyChangedCallback(AddressOf UpdateColorRGB), New CoerceValueCallback(AddressOf RGBCoerce)))


    Private Shared Function RGBCoerce(ByVal d As DependencyObject, ByVal value As Object) As Object
        Dim v As Integer = CInt(Fix(value))
        If v < 0 Then
            Return 0
        End If
        If v > 255 Then
            Return 255
        End If
        Return v
    End Function

#End Region

    ''' <summary>
    ''' Shared property changed callback to update the Color property
    ''' </summary>
    Public Shared Sub UpdateColorHSB(ByVal o As Object, ByVal e As DependencyPropertyChangedEventArgs)
        Dim c As ColorPicker = CType(o, ColorPicker)
        If c._HSBSetInternally Then Return

        c._RGBSetInternally = True
        Dim n As Color = ColorHelper.ColorFromAHSB(c.Alpha, c.Hue, c.Saturation, c.Brightness)
        c.Color = n

        If c.SelectedGradient IsNot Nothing Then
            c.SelectedGradient.Color = n
        End If

        c.SetBrush()
        c._HSBSetInternally = False
    End Sub

    ''' <summary>
    ''' Shared property changed callback to update the Color property
    ''' </summary>
    Public Shared Sub UpdateColorRGB(ByVal o As Object, ByVal e As DependencyPropertyChangedEventArgs)
        Dim c As ColorPicker = CType(o, ColorPicker)
        If c._RGBSetInternally Then Return

        c._RGBSetInternally = True
        Dim n As Color = Color.FromArgb(
            CByte(c.A),
            CByte(c.R),
            CByte(c.G),
            CByte(c.B)
        )
        c.Color = n

        If c.SelectedGradient IsNot Nothing Then
            c.SelectedGradient.Color = n
        End If

        c.SetBrush()
        c._RGBSetInternally = False
    End Sub


#Region "ColorChanged Event"

    Public Delegate Sub ColorChangedEventHandler(ByVal sender As Object, ByVal e As ColorChangedEventArgs)

    Public Shared ReadOnly ColorChangedEvent As RoutedEvent = EventManager.RegisterRoutedEvent("ColorChanged", RoutingStrategy.Bubble, GetType(ColorChangedEventHandler), GetType(ColorPicker))

    Public Custom Event ColorChanged As ColorChangedEventHandler
        AddHandler(ByVal value As ColorChangedEventHandler)
            MyBase.AddHandler(ColorChangedEvent, value)
        End AddHandler
        RemoveHandler(ByVal value As ColorChangedEventHandler)
            MyBase.RemoveHandler(ColorChangedEvent, value)
        End RemoveHandler
        RaiseEvent(ByVal sender As Object, ByVal e As ColorChangedEventArgs)
            MyBase.RaiseEvent(e)
        End RaiseEvent
    End Event

    Private Sub RaiseColorChangedEvent(ByVal color As Color)
        Dim newEventArgs As New ColorChangedEventArgs(ColorPicker.ColorChangedEvent, color)
        MyBase.RaiseEvent(newEventArgs)
    End Sub

#End Region


    Friend Sub SetBrush(Optional BrushTypeChanged As Boolean = False)
        If Not _UpdateBrush Then Return

        Me._BrushSetInternally = Not BrushTypeChanged

        ' retain old opacity
        Dim _opacity As Double = 1
        If Me.Brush IsNot Nothing Then
            _opacity = Me.Brush.Opacity
        End If

        Select Case BrushType
            Case BrushTypes.None
                Brush = Nothing

            Case BrushTypes.Solid
                Brush = New SolidColorBrush(Me.Color)

            Case BrushTypes.Linear
                Dim _brush = New LinearGradientBrush()
                For Each _g As GradientStop In Gradients
                    _brush.GradientStops.Add(New GradientStop(_g.Color, _g.Offset))
                Next
                _brush.StartPoint = New Point(Me.StartX, Me.StartY)
                _brush.EndPoint = New Point(Me.EndX, Me.EndY)
                _brush.ColorInterpolationMode = Me.ColorInterpolationMode
                _brush.SpreadMethod = Me.SpreadMethod
                Brush = _brush

            Case BrushTypes.Radial
                Dim brush1 = New RadialGradientBrush()
                For Each _g As GradientStop In Gradients
                    brush1.GradientStops.Add(New GradientStop(_g.Color, _g.Offset))
                Next

                brush1.GradientOrigin = New Point(Me.GradientOriginX, Me.GradientOriginY)
                brush1.Center = New Point(Me.CenterX, Me.CenterY)
                brush1.RadiusX = Me.RadiusX
                brush1.RadiusY = Me.RadiusY
                brush1.ColorInterpolationMode = Me.ColorInterpolationMode
                brush1.SpreadMethod = Me.SpreadMethod
                Brush = brush1
        End Select

        If Me.BrushType <> BrushTypes.None Then
            Me.Brush.Opacity = _opacity ' retain old opacity
            If Brush.RelativeTransform IsNot TransPker.Transform Then
                Brush.RelativeTransform = TransPker.Transform
            End If
        End If

        Me._BrushSetInternally = False
        Me.RaiseColorChangedEvent(Me.Color)
    End Sub

    Private Sub Scv_ScrollChanged(sender As Object, e As ScrollChangedEventArgs)
        Static LastViewportWidth As Double = 0

        Dim StkPnl As StackPanel = Scv.Content
        Dim MinVal = Me.MinWidth - If(Scv.ComputedVerticalScrollBarVisibility = Windows.Visibility.Visible, SystemParameters.ScrollWidth, 0)
        StkPnl.Width = Math.Max(MinVal, Scv.ViewportWidth - 10)
        If LastViewportWidth <> e.ViewportWidth Then
            LastViewportWidth = e.ViewportWidth
            ColorLst.ComputeSize()
        End If

    End Sub

    Private Sub TransPker_TransformChanged(Tp As TransformPicker, OldTransform As Transform, NewTransform As Transform)
        Me.Brush.RelativeTransform = Tp.Transform
    End Sub

    Private Sub ColorPicker_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        If Me.Brush Is Nothing Then Return
        Dim T = Me.Brush.GetType
        If T Is GetType(LinearGradientBrush) OrElse T Is GetType(RadialGradientBrush) Then
            TransPker.Transform = Me.Brush.RelativeTransform
        End If
    End Sub
End Class