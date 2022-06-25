Imports System.Collections.Generic
Imports Microsoft.Nautilus.Core

Namespace Microsoft.Nautilus.Text
    Public NotInheritable Class ContentTypeAttribute
        Inherits BaseMetadataProviderAttribute

        Private name As List(Of String)

        Public ReadOnly Property ContentTypes As IEnumerable(Of String)
            Get
                Return name
            End Get
        End Property

        Public Sub New(name As String)
            If name Is Nothing OrElse name.Length = 0 Then
                Throw New ArgumentNullException("name")
            End If

            Me.name = New List(Of String)
            Me.name.Add(name)
        End Sub
    End Class
End Namespace
