Imports System.Collections.Generic

Namespace Microsoft.Nautilus.Core
    Public Class OrderAfterAttribute
        Inherits BaseMetadataProviderAttribute

        Private _after As List(Of String)

        Public ReadOnly Property After As IEnumerable(Of String)
            Get
                Return _after
            End Get
        End Property

        Public Sub New(after As String)
            _after = New List(Of String)
            _after.Add(after)
        End Sub
    End Class
End Namespace
