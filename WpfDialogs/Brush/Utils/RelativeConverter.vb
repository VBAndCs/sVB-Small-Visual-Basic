Imports System.Globalization

Friend Class RelativeConverter
    Implements IValueConverter

    Public RelativeTo As Double = 1

    Public Function Convert(ByVal value As Object, ByVal targetType As Type, ByVal parameter As Object, ByVal culture As CultureInfo) As Object Implements IValueConverter.Convert
        If value < 1 Then Return value
        Return value / RelativeTo
    End Function

    Public Function ConvertBack(ByVal value As Object, ByVal targetType As Type, ByVal parameter As Object, ByVal culture As CultureInfo) As Object Implements IValueConverter.ConvertBack
        Return value * RelativeTo
    End Function
End Class