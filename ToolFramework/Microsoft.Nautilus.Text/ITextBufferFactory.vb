Imports System.IO
Imports Microsoft.Nautilus.Core.Task

Namespace Microsoft.Nautilus.Text
    Public Interface ITextBufferFactory
        Inherits IBufferTracker
        Function CreateTextBuffer() As ITextBuffer

        Function CreateTextBuffer(text As String) As ITextBuffer

        Function CreateTextBuffer(text As String, contentType As String) As ITextBuffer

        Function CreateTextBuffer(reader As TextReader) As ITextBuffer

        Function CreateTextBuffer(reader As TextReader, contentType As String) As ITextBuffer

        Function CreateTextBuffer(reader As TextReader, contentType As String, cancelableTask1 As CancelableTask) As ITextBuffer
    End Interface
End Namespace
