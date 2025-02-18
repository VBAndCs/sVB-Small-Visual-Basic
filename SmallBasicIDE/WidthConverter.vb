Imports System
Imports System.Globalization
Imports System.Windows
Imports System.Windows.Data

Namespace Microsoft.SmallVisualBasic
    Public Class WidthConverter
        Implements IMultiValueConverter

        Public Function Convert(values As Object(), targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IMultiValueConverter.Convert
            If values.Length = 3 AndAlso TypeOf values(0) Is Double AndAlso TypeOf values(1) Is Double AndAlso TypeOf values(2) Is Thickness Then
                Dim totalWidth As Double = DirectCast(values(0), Double)
                Dim arrowWidth As Double = DirectCast(values(1), Double)
                Dim padding As Thickness = DirectCast(values(2), Thickness)
                Return Math.Max(totalWidth - arrowWidth - padding.Left - padding.Right * 2, 0)
            End If
            Return 0
        End Function

        Public Function ConvertBack(value As Object, targetTypes As Type(), parameter As Object, culture As CultureInfo) As Object() Implements IMultiValueConverter.ConvertBack
            Throw New NotImplementedException()
        End Function
    End Class
End Namespace
