Imports System.Globalization

Friend Class CmConverter
    Implements IValueConverter

    Public Function Convert(ByVal value As Object, ByVal targetType As Type, ByVal parameter As Object, ByVal culture As CultureInfo) As Object Implements IValueConverter.Convert
        Dim d As Double = CDbl(value) * Helper.PxToCm
        Return Math.Round(d, 2)
    End Function

    Public Function ConvertBack(ByVal value As Object, ByVal targetType As Type, ByVal parameter As Object, ByVal culture As CultureInfo) As Object Implements IValueConverter.ConvertBack
        Return Nothing
    End Function
End Class

