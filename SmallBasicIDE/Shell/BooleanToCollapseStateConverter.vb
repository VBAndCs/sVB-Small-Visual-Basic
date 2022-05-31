Imports System
Imports System.Globalization
Imports System.Windows
Imports System.Windows.Data

Namespace Microsoft.SmallBasic.Shell
    Public Class BooleanToCollapseStateConverter
        Implements IValueConverter

        Public Function Convert(
                               value As Object,
                               targetType As Type,
                               parameter As Object,
                               culture As CultureInfo
                  ) As Object Implements IValueConverter.Convert

            If targetType Is GetType(Visibility) AndAlso TypeOf value Is Boolean Then
                If value Then
                    Return Visibility.Collapsed
                End If

                Return Visibility.Visible
            End If

            Throw New InvalidOperationException()
        End Function

        Public Function ConvertBack(ByVal value As Object, ByVal targetType As Type, ByVal parameter As Object, ByVal culture As CultureInfo) As Object Implements IValueConverter.ConvertBack
            Throw New NotImplementedException()
        End Function
    End Class
End Namespace
