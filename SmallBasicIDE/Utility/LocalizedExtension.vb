Imports System
Imports System.Windows.Markup
Imports Microsoft.SmallBasic

Namespace Microsoft.SmallVisualBasic.Utility
    <MarkupExtensionReturnType(GetType(String))>
    Public Class LocalizedExtension
        Inherits MarkupExtension
        Public Property ResourceId As String

        Public Sub New(resourceId As String)
            Me.ResourceId = resourceId
        End Sub

        Public Overrides Function ProvideValue(serviceProvider As IServiceProvider) As Object
            Return ResourceHelper.GetString(ResourceId)
        End Function
    End Class
End Namespace
