Imports System.IO

Namespace Microsoft.Nautilus.Text
    Public Interface IBufferTracker
        Property TrackBuffers As Boolean

        Sub TagBuffer(buffer As ITextBuffer, tag As String)

        Function ReportLiveBuffers(writer As TextWriter) As Integer
    End Interface
End Namespace
