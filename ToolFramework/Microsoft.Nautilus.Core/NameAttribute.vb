Namespace Microsoft.Nautilus.Core
    Public Class NameAttribute
        Inherits BaseMetadataProviderAttribute

        Public ReadOnly Property Name As String

        Public Sub New(name As String)
            _Name = name
        End Sub
    End Class
End Namespace
