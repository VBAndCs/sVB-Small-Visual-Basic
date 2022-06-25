Imports System.Windows.Markup
Imports Microsoft.Nautilus.Text

Namespace Microsoft.Windows.Controls
    Public Class CreateTextBufferExtension
        Inherits MarkupExtension

        Public Property ContentType As String

        Public Overrides Function ProvideValue(serviceProvider As IServiceProvider) As Object
            Return New BufferFactory().CreateTextBuffer("", ContentType)
        End Function
    End Class
End Namespace
