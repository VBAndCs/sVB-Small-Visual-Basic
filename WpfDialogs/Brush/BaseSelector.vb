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


Friend MustInherit Class BaseSelector
    Inherits FrameworkElement

    Public Property Orientation() As Orientation
        Get
            Return CType(GetValue(OrientationProperty), Orientation)
        End Get
        Set(ByVal value As Orientation)
            SetValue(OrientationProperty, value)
        End Set
    End Property
    Public Shared ReadOnly OrientationProperty As DependencyProperty = DependencyProperty.Register("Orientation", GetType(Orientation), GetType(BaseSelector), New UIPropertyMetadata(Orientation.Vertical))

    Public Property Color() As Color
        Get
            Return CType(GetValue(ColorProperty), Color)
        End Get
        Protected Set(ByVal value As Color)
            SetValue(ColorProperty, value)
        End Set
    End Property
    Public Shared ReadOnly ColorProperty As DependencyProperty = DependencyProperty.Register("Color", GetType(Color), GetType(BaseSelector), New UIPropertyMetadata(Colors.Red))
End Class