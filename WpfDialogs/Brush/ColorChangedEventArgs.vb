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


Public Class ColorChangedEventArgs
    Inherits RoutedEventArgs

    Public Sub New(ByVal routedEvent As RoutedEvent, ByVal color As Color)
        Me.RoutedEvent = routedEvent
        Me.Color = color
    End Sub

    Private _Color As Color
    Public Property Color() As Color
        Get
            Return _Color
        End Get
        Set(ByVal value As Color)
            _Color = value
        End Set
    End Property
End Class