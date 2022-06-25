Imports System.IO
Imports System.Text

Namespace Microsoft.Nautilus.Text.Document
    Public Interface IEncodingSniffer
        Function GetStreamEncoding(stream1 As Stream) As Encoding
    End Interface
End Namespace
