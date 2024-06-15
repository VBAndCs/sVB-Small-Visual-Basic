Public Class SBControl
    Inherits Image

    Public Property Control As Control

    Dim _displayImage As String
    Public Property DisplayImage As String
        Get
            Return _displayImage
        End Get

        Set(value As String)
            _displayImage = IO.Path.GetFileName(value)
            Dim appPath = System.AppDomain.CurrentDomain.BaseDirectory
            Dim imgPath = IO.Path.Combine(appPath, "ToolBox\Controls\" & _displayImage)
            Me.Source = New BitmapImage(New Uri(imgPath, UriKind.Absolute))
        End Set
    End Property
End Class
