Friend Class SaturationBrightnessSelector
    Inherits BaseSelector

    Public Property OffsetPadding() As Thickness
        Get
            Return CType(GetValue(OffsetPaddingProperty), Thickness)
        End Get
        Set(ByVal value As Thickness)
            SetValue(OffsetPaddingProperty, value)
        End Set
    End Property
    Public Shared ReadOnly OffsetPaddingProperty As DependencyProperty = DependencyProperty.Register("OffsetPadding", GetType(Thickness), GetType(SaturationBrightnessSelector), New UIPropertyMetadata(New Thickness(0.0)))


    Public Property Hue() As Double
        Private Get
            Return CDbl(GetValue(HueProperty))
        End Get
        Set(ByVal value As Double)
            SetValue(HueProperty, value)
        End Set
    End Property

    Public Shared ReadOnly HueProperty As DependencyProperty = DependencyProperty.Register("Hue", GetType(Double), GetType(SaturationBrightnessSelector), New FrameworkPropertyMetadata(0.0, FrameworkPropertyMetadataOptions.AffectsRender, New PropertyChangedCallback(AddressOf HueChanged)))
    Public Shared Sub HueChanged(ByVal o As Object, ByVal e As DependencyPropertyChangedEventArgs)
        Dim h As SaturationBrightnessSelector = CType(o, SaturationBrightnessSelector)
        h.SetColor()
    End Sub


    Public Property SaturationOffset() As Double
        Get
            Return CDbl(GetValue(SaturationOffsetProperty))
        End Get
        Set(ByVal value As Double)
            SetValue(SaturationOffsetProperty, value)
        End Set
    End Property

    Public Shared ReadOnly SaturationOffsetProperty As DependencyProperty = DependencyProperty.Register("SaturationOffset", GetType(Double), GetType(SaturationBrightnessSelector), New UIPropertyMetadata(0.0))


    Public Property Saturation() As Double
        Get
            Return CDbl(GetValue(SaturationProperty))
        End Get
        Set(ByVal value As Double)
            SetValue(SaturationProperty, value)
        End Set
    End Property

    Public Shared ReadOnly SaturationProperty As DependencyProperty = DependencyProperty.Register("Saturation", GetType(Double), GetType(SaturationBrightnessSelector), New FrameworkPropertyMetadata(0.0, New PropertyChangedCallback(AddressOf SaturationChanged), New CoerceValueCallback(AddressOf SaturationCoerce)))
    Public Shared Sub SaturationChanged(ByVal o As Object, ByVal e As DependencyPropertyChangedEventArgs)
        Dim h As SaturationBrightnessSelector = CType(o, SaturationBrightnessSelector)
        h.SetSaturationOffset()
    End Sub
    Public Shared Function SaturationCoerce(ByVal d As DependencyObject, ByVal Brightness As Object) As Object
        Dim v As Double = CDbl(Brightness)
        If v < 0 Then
            Return 0.0
        End If
        If v > 1 Then
            Return 1.0
        End If
        Return v
    End Function

    Public Property BrightnessOffset() As Double
        Get
            Return CDbl(GetValue(BrightnessOffsetProperty))
        End Get
        Set(ByVal value As Double)
            SetValue(BrightnessOffsetProperty, value)
        End Set
    End Property

    Public Shared ReadOnly BrightnessOffsetProperty As DependencyProperty = DependencyProperty.Register("BrightnessOffset", GetType(Double), GetType(SaturationBrightnessSelector), New UIPropertyMetadata(0.0))


    Public Property Brightness() As Double
        Get
            Return CDbl(GetValue(BrightnessProperty))
        End Get
        Set(ByVal value As Double)
            SetValue(BrightnessProperty, value)
        End Set
    End Property

    Public Shared ReadOnly BrightnessProperty As DependencyProperty = DependencyProperty.Register("Brightness", GetType(Double), GetType(SaturationBrightnessSelector), New FrameworkPropertyMetadata(0.0, New PropertyChangedCallback(AddressOf BrightnessChanged), New CoerceValueCallback(AddressOf BrightnessCoerce)))

    Public Shared Sub BrightnessChanged(ByVal o As Object, ByVal e As DependencyPropertyChangedEventArgs)
        Dim h As SaturationBrightnessSelector = CType(o, SaturationBrightnessSelector)
        h.SetBrightnessOffset()
    End Sub

    Public Shared Function BrightnessCoerce(ByVal d As DependencyObject, ByVal Brightness As Object) As Object
        Dim v As Double = CDbl(Brightness)
        If v < 0 Then
            Return 0.0
        End If
        If v > 1 Then
            Return 1.0
        End If
        Return v
    End Function


    Protected Overrides Sub OnMouseMove(ByVal e As System.Windows.Input.MouseEventArgs)
        If e.LeftButton = MouseButtonState.Pressed Then
            Dim p As Point = e.GetPosition(Me)
            ColorPicker._HSBSetInternally = True
            Saturation = (p.X / (Me.ActualWidth - OffsetPadding.Right))
            ColorPicker._HSBSetInternally = True
            Brightness = (Me.ActualHeight - OffsetPadding.Bottom - p.Y) / (Me.ActualHeight - OffsetPadding.Bottom)
            ColorPicker._HSBSetInternally = False
            SetColor()
        End If
        MyBase.OnMouseMove(e)
    End Sub

    Function GetParent(Child As UIElement, ParentType As Type)
        Dim P = Child
        Do
            P = VisualTreeHelper.GetParent(P)
            If P Is Nothing OrElse P.GetType Is ParentType Then Return P
        Loop
    End Function

    Protected Overrides Sub OnMouseDown(ByVal e As MouseButtonEventArgs)
        Dim p As Point = e.GetPosition(Me)
        ColorPicker._HSBSetInternally = True
        Saturation = (p.X / (Me.ActualWidth - OffsetPadding.Right))
        ColorPicker._HSBSetInternally = True
        Brightness = (Me.ActualHeight - OffsetPadding.Bottom - p.Y) / (Me.ActualHeight - OffsetPadding.Bottom)
        ColorPicker._HSBSetInternally = False
        SetColor()

        Mouse.Capture(Me)
        MyBase.OnMouseDown(e)
    End Sub

    Protected Overrides Sub OnMouseUp(ByVal e As MouseButtonEventArgs)
        Me.ReleaseMouseCapture()
        MyBase.OnMouseUp(e)
    End Sub

    Protected Overrides Sub OnRender(ByVal dc As DrawingContext)
        Dim h As New LinearGradientBrush()
        h.StartPoint = New Point(0, 0)
        h.EndPoint = New Point(1, 0)
        h.GradientStops.Add(New GradientStop(Colors.White, 0.0))
        h.GradientStops.Add(New GradientStop(ColorHelper.ColorFromHSB(Hue, 1, 1), 1.0))
        dc.DrawRectangle(h, Nothing, New Rect(0, 0, ActualWidth, ActualHeight))

        Dim v As New LinearGradientBrush()
        v.StartPoint = New Point(0, 0)
        v.EndPoint = New Point(0, 1)
        v.GradientStops.Add(New GradientStop(Color.FromArgb(&HFF, 0, 0, 0), 1.0))
        v.GradientStops.Add(New GradientStop(Color.FromArgb(&H80, 0, 0, 0), 0.5))
        v.GradientStops.Add(New GradientStop(Color.FromArgb(&H0, 0, 0, 0), 0.0))
        dc.DrawRectangle(v, Nothing, New Rect(0, 0, ActualWidth, ActualHeight))

        SetSaturationOffset()
        SetBrightnessOffset()
    End Sub

    Private Sub SetSaturationOffset()
        SaturationOffset = OffsetPadding.Left + ((ActualWidth - (OffsetPadding.Right + OffsetPadding.Left)) * Saturation)
    End Sub

    Private Sub SetBrightnessOffset()
        BrightnessOffset = OffsetPadding.Top + ((ActualHeight - (OffsetPadding.Bottom + OffsetPadding.Top)) - ((ActualHeight - (OffsetPadding.Bottom + OffsetPadding.Top)) * Brightness))
    End Sub

    Public Sub SetColor()
        Color = ColorHelper.ColorFromHSB(Hue, Saturation, Brightness)
        If ColorPicker IsNot Nothing Then ColorPicker.Color = Color
    End Sub

    Dim ColorPicker As ColorPicker
    Private Sub SaturationBrightnessSelector_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        ColorPicker = GetParent(Me, GetType(ColorPicker))
    End Sub
End Class
