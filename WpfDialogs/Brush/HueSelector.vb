'''***************   NCore Softwares Pvt. Ltd., India   **************************
'''
'''   ColorPicker
'''
'''   Copyright (C) 2013 NCore Softwares Pvt. Ltd.
'''
'''   This program is provided to you under the terms of the Microsoft Public
'''   License (Ms-PL) as published at http://ColorPicker.codeplex.com/license
'''
'''**********************************************************************************


Friend Class HueSelector
    Inherits BaseSelector

    Public Property Hue() As Double
        Get
            Return CDbl(GetValue(HueProperty))
        End Get
        Set(ByVal value As Double)
            SetValue(HueProperty, value)
        End Set
    End Property

    Public Shared ReadOnly HueProperty As DependencyProperty = DependencyProperty.Register("Hue", GetType(Double), GetType(HueSelector), New FrameworkPropertyMetadata(0.0, New PropertyChangedCallback(AddressOf HueChanged), New CoerceValueCallback(AddressOf HueCoerce)))

    Public Shared Sub HueChanged(ByVal o As Object, ByVal e As DependencyPropertyChangedEventArgs)
        Dim h As HueSelector = CType(o, HueSelector)
        h.SetHueOffset()
        h.SetColor()
    End Sub

    Public Shared Function HueCoerce(ByVal d As DependencyObject, ByVal Brightness As Object) As Object
        Dim v As Double = CDbl(Brightness)
        If v < 0 Then
            Return 0.0
        End If
        If v > 1 Then
            Return 1.0
        End If
        Return v
    End Function

    Public Property HueOffset() As Double
        Get
            Return CDbl(GetValue(HueOffsetProperty))
        End Get
        Private Set(ByVal value As Double)
            SetValue(HueOffsetProperty, value)
        End Set
    End Property

    Public Shared ReadOnly HueOffsetProperty As DependencyProperty = DependencyProperty.Register("HueOffset", GetType(Double), GetType(HueSelector), New UIPropertyMetadata(0.0))


    Protected Overrides Sub OnMouseMove(ByVal e As MouseEventArgs)
        If e.LeftButton = MouseButtonState.Pressed Then
            Dim p As Point = e.GetPosition(Me)

            If Orientation = Orientation.Vertical Then
                Hue = 1 - (p.Y / Me.ActualHeight)
            Else
                Hue = 1 - (p.X / Me.ActualWidth)
            End If
        End If
        MyBase.OnMouseMove(e)
    End Sub

    Protected Overrides Sub OnMouseDown(ByVal e As MouseButtonEventArgs)
        If e.LeftButton = MouseButtonState.Pressed Then
            Dim p As Point = e.GetPosition(Me)

            If Orientation = Orientation.Vertical Then
                Hue = 1 - (p.Y / Me.ActualHeight)
            Else
                Hue = 1 - (p.X / Me.ActualWidth)
            End If
        End If
        Mouse.Capture(Me)
        MyBase.OnMouseMove(e)
    End Sub

    Protected Overrides Sub OnMouseUp(ByVal e As MouseButtonEventArgs)
        Me.ReleaseMouseCapture()
        MyBase.OnMouseUp(e)
    End Sub

    Protected Overrides Sub OnRender(ByVal dc As DrawingContext)
        Dim lb As New LinearGradientBrush()

        lb.StartPoint = New Point(0, 0)

        If Orientation = Orientation.Vertical Then
            lb.EndPoint = New Point(0, 1)
        Else
            lb.EndPoint = New Point(1, 0)
        End If

        lb.GradientStops.Add(New GradientStop(Color.FromRgb(&HFF, &H0, &H0), 1.0))
        lb.GradientStops.Add(New GradientStop(Color.FromRgb(&HFF, &HFF, &H0), 0.85))
        lb.GradientStops.Add(New GradientStop(Color.FromRgb(&H0, &HFF, &H0), 0.76))
        lb.GradientStops.Add(New GradientStop(Color.FromRgb(&H0, &HFF, &HFF), 0.5))
        lb.GradientStops.Add(New GradientStop(Color.FromRgb(&H0, &H0, &HFF), 0.33))
        lb.GradientStops.Add(New GradientStop(Color.FromRgb(&HFF, &H0, &HFF), 0.16))
        lb.GradientStops.Add(New GradientStop(Color.FromRgb(&HFF, &H0, &H0), 0.0))

        dc.DrawRectangle(lb, Nothing, New Rect(0, 0, ActualWidth, ActualHeight))

        SetHueOffset()
    End Sub

    Protected Overrides Function ArrangeOverride(ByVal finalSize As Size) As Size
        SetHueOffset()
        Return MyBase.ArrangeOverride(finalSize)
    End Function


    Private Sub SetHueOffset()
        Dim length As Double = ActualHeight
        If Orientation = Orientation.Horizontal Then
            length = ActualWidth
        End If

        HueOffset = length - (length * Hue)
    End Sub

    Private Sub SetColor()
        MyBase.Color = ColorHelper.ColorFromHSB(Hue, 1, 1)
        'base.Brush = new SolidColorBrush(Color);
    End Sub
End Class
