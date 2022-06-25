Imports Microsoft.Nautilus.Core

Namespace Microsoft.Nautilus.Text.AdornmentSystem
    Public NotInheritable Class AdornmentDiscriminatorAttribute
        Inherits BaseMetadataProviderAttribute

        Public ReadOnly Property SupportedAdornmentType As String

        Public Sub New(supportedAdornmentType As Type)
            If supportedAdornmentType Is Nothing Then
                Throw New ArgumentNullException("supportedAdornmentType")
            End If

            _SupportedAdornmentType = supportedAdornmentType.AssemblyQualifiedName
        End Sub
    End Class
End Namespace
