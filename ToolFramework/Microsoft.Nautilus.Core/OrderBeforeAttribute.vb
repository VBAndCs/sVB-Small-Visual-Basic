Imports System.Collections.Generic

Namespace Microsoft.Nautilus.Core
    Public Class OrderBeforeAttribute
        Inherits BaseMetadataProviderAttribute

        Private _before As List(Of String)

        Public ReadOnly Property Before As IEnumerable(Of String)
            Get
                Return _before
            End Get
        End Property

        Public Sub New(before As String)
            _before = New List(Of String)
            _before.Add(before)
        End Sub
    End Class
End Namespace
