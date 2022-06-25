
Namespace Microsoft.Nautilus.Core
    Public Class BaseTypeAttribute
        Inherits BaseMetadataProviderAttribute

        Public ReadOnly Property BaseType As String


        Public Sub New(baseType As String)
            If String.IsNullOrEmpty(baseType) Then
                Throw New ArgumentException("baseType")
            End If

            _BaseType = baseType
        End Sub
    End Class
End Namespace
