Public Class ImageBrushes
    Inherits DependencyObject

    Public Shared Function GetImageFileName(ByVal element As DependencyObject) As String
        If element Is Nothing Then
            Throw New ArgumentNullException("element")
        End If

        Return element.GetValue(ImageFileNameProperty)
    End Function

    Public Shared Sub SetImageFileName(ByVal element As DependencyObject, ByVal value As String)
        If element Is Nothing Then
            Throw New ArgumentNullException("element")
        End If

        element.SetValue(ImageFileNameProperty, value)
    End Sub

    Public Shared ReadOnly ImageFileNameProperty As  _
                           DependencyProperty = DependencyProperty.RegisterAttached("ImageFileName", _
                           GetType(String), GetType(ImageBrushes), _
                           New PropertyMetadata(Nothing))



End Class
