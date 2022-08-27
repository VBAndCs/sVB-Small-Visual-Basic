Imports System
Imports System.Globalization
Imports System.Windows
Imports System.Windows.Data

Namespace Microsoft.SmallBasic.Shell
    Public Class NullToVisibilityConverter
        Implements IValueConverter

        Public Function Convert(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.Convert
            If targetType Is GetType(Visibility) Then
                If value IsNot Nothing Then
                    Return Visibility.Visible
                End If

                Return Visibility.Collapsed
            End If

            Throw New InvalidOperationException()
        End Function

        Public Function ConvertBack(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.ConvertBack
            Throw New NotImplementedException()
        End Function
    End Class
End Namespace
