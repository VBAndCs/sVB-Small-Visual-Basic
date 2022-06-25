Imports Microsoft.Nautilus.Core

Namespace Microsoft.Nautilus.Text.Adornments
    Public NotInheritable Class VisualFactoryTypeAttribute
        Inherits BaseMetadataProviderAttribute

        Public ReadOnly Property VisualFactoryType As String

        Public Sub New(type As String)
            If type Is Nothing Then
                Throw New ArgumentNullException("type")
            End If

            _VisualFactoryType = type
        End Sub
    End Class
End Namespace
