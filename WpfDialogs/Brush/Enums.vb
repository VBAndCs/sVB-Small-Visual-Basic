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


Public Enum SpinDirection
    Increase
    Decrease
End Enum


Friend Enum ValidSpinDirections
    None
    Increase
    Decrease
End Enum


Friend Enum AllowedSpecialValues
    None = 0
    NaN = 1
    PositiveInfinity = 2
    NegativeInfinity = 4
    AnyInfinity = PositiveInfinity Or NegativeInfinity
    Any = NaN Or AnyInfinity
End Enum


Friend Enum BrushTypes
    None
    Solid
    Linear
    Radial
End Enum