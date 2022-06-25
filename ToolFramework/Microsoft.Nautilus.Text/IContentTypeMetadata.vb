Imports System.Collections.Generic

Namespace Microsoft.Nautilus.Text
    Public Interface IContentTypeMetadata
        ReadOnly Property ContentTypes As IEnumerable(Of String)
    End Interface
End Namespace
