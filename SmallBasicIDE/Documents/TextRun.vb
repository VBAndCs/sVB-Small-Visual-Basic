﻿Imports System.Windows
Imports System.Windows.Documents
Imports System.Windows.Threading

Namespace Microsoft.SmallVisualBasic.Utility
    Public Class TextRun
        Inherits Run

        Public Shared Shadows ReadOnly TextProperty As DependencyProperty =
            DependencyProperty.Register(
                      "Text",
                      GetType(String),
                      GetType(TextRun),
                      New PropertyMetadata(AddressOf OnTextChanged)
            )

        Public Overloads Property Text As String
            Get
                Return CStr(GetValue(TextProperty))
            End Get
            Set(value As String)
                SetValue(TextProperty, value)
            End Set
        End Property

        Private Shared Sub OnTextChanged(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
            CType(d, Run).Dispatcher.BeginInvoke(
                   DispatcherPriority.Normal,
                   CType(Function()
                             CType(d, Run).Text = CStr(e.NewValue)
                             Return Nothing
                         End Function,
                         DispatcherOperationCallback),
                   Nothing)
        End Sub
    End Class
End Namespace
