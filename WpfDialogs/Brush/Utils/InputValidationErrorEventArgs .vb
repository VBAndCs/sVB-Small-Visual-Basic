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

Imports System.Text

Public Delegate Sub InputValidationErrorEventHandler(ByVal sender As Object, ByVal e As InputValidationErrorEventArgs)

Public Class InputValidationErrorEventArgs
    Inherits EventArgs

    Private _exception As Exception
    Private _throwException As Boolean

    Public Sub New(ByVal e As Exception)
        Exception = e
    End Sub


    Public Property Exception() As Exception
        Get
            Return _exception
        End Get
        Private Set(ByVal value As Exception)
            _exception = value
        End Set
    End Property

    Public Property ThrowException() As Boolean
        Get
            Return _throwException
        End Get
        Set(ByVal value As Boolean)
            _throwException = value
        End Set
    End Property
End Class

Friend Interface IValidateInput
    Event InputValidationError As InputValidationErrorEventHandler
    Function CommitInput() As Boolean
End Interface