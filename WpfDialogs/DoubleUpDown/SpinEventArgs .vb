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


Public Class SpinEventArgs
    Inherits RoutedEventArgs

    Private privateDirection As SpinDirection
    Public Property Direction() As SpinDirection
        Get
            Return privateDirection
        End Get
        Private Set(ByVal value As SpinDirection)
            privateDirection = value
        End Set
    End Property

    Public Sub New(ByVal direction As SpinDirection)
        MyBase.New()
        Me.Direction = direction
    End Sub
End Class
