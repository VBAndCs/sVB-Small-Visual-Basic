Imports System.IO
Imports System.Text

Namespace Microsoft.Nautilus.Text.Document
    Public Interface IEncodingManager
        Function GetSniffedEncoding(stream1 As Stream) As Encoding
    End Interface
End Namespace
