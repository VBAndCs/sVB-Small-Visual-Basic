Imports System.ComponentModel.Composition

Namespace Microsoft.Nautilus.Text.Projection.Implementation
    <Export(GetType(IBufferGraphFactory))>
    Public NotInheritable Class BufferGraphFactory
        Implements IBufferGraphFactory

        Public Function CreateBufferGraph(textBuffer As ITextBuffer) As IBufferGraph Implements IBufferGraphFactory.CreateBufferGraph
            Return New BufferGraph(textBuffer)
        End Function
    End Class
End Namespace
