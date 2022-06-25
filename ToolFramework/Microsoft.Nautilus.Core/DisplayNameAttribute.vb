
Namespace Microsoft.Nautilus.Core
    Public Class DisplayNameAttribute
        Inherits BaseMetadataProviderAttribute

        Public ReadOnly Property DisplayName As String

        Public Sub New(displayName As String)
            If displayName Is Nothing Then
                Throw New ArgumentNullException("displayName")
            End If
            _DisplayName = displayName
        End Sub
    End Class
End Namespace
