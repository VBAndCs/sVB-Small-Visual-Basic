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


Friend Class GradientStopSlider
    Inherits Slider

    Protected Overrides Sub OnPreviewMouseLeftButtonDown(ByVal e As MouseButtonEventArgs)
        MyBase.OnPreviewMouseLeftButtonDown(e)

        If Me.ColorPicker IsNot Nothing Then
            Me.ColorPicker._BrushSetInternally = True
            Me.ColorPicker._UpdateBrush = False

            Me.ColorPicker.SelectedGradient = Me.SelectedGradient
            Me.ColorPicker.Color = Me.SelectedGradient.Color

            Me.ColorPicker._UpdateBrush = True
            'this.ColorPicker._BrushSetInternally = false;

            'e.Handled = true;
        End If
    End Sub

    Protected Overrides Sub OnValueChanged(ByVal oldValue As Double, ByVal newValue As Double)
        MyBase.OnValueChanged(oldValue, newValue)

        If Me.ColorPicker IsNot Nothing Then
            'this.ColorPicker._HSBSetInternally = true;
            'this.ColorPicker._RGBSetInternally = true;
            Me.ColorPicker._BrushSetInternally = True
            Me.ColorPicker.SetBrush()
            Me.ColorPicker._HSBSetInternally = False
            'this.ColorPicker._RGBSetInternally = false;
            'this.ColorPicker._BrushSetInternally = false;
        End If
    End Sub

    Public Property ColorPicker() As ColorPicker
        Get
            Return CType(GetValue(ColorPickerProperty), ColorPicker)
        End Get
        Set(ByVal value As ColorPicker)
            SetValue(ColorPickerProperty, value)
        End Set
    End Property
    Public Shared ReadOnly ColorPickerProperty As DependencyProperty = DependencyProperty.Register("ColorPicker", GetType(ColorPicker), GetType(GradientStopSlider))

    Public Property SelectedGradient() As GradientStop
        Get
            Return CType(GetValue(SelectedGradientProperty), GradientStop)
        End Get
        Set(ByVal value As GradientStop)
            SetValue(SelectedGradientProperty, value)
        End Set
    End Property
    Public Shared ReadOnly SelectedGradientProperty As DependencyProperty = DependencyProperty.Register("SelectedGradient", GetType(GradientStop), GetType(GradientStopSlider))
End Class