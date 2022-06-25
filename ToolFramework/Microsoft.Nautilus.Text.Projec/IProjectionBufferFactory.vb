Imports System.Collections.Generic

Namespace Microsoft.Nautilus.Text.Projection
    Public Interface IProjectionBufferFactory
        Inherits IBufferTracker
        Function CreateProjectionBuffer(projectionEditResolver As IProjectionEditResolver, contentType As String, textSpans As IList(Of ITextSpan)) As IProjectionBuffer
    End Interface
End Namespace
