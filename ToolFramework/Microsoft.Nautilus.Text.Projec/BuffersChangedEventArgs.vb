Imports System.Collections.Generic
Imports System.Collections.ObjectModel

Namespace Microsoft.Nautilus.Text.Projection
    Public Class BuffersChangedEventArgs
        Inherits EventArgs

        Private _addedBuffers As IList(Of ITextBuffer)
        Private _removedBuffers As IList(Of ITextBuffer)

        Public ReadOnly Property AddedBuffers As ReadOnlyCollection(Of ITextBuffer)
            Get
                Return New ReadOnlyCollection(Of ITextBuffer)(_addedBuffers)
            End Get
        End Property

        Public ReadOnly Property RemovedBuffers As ReadOnlyCollection(Of ITextBuffer)
            Get
                Return New ReadOnlyCollection(Of ITextBuffer)(_removedBuffers)
            End Get
        End Property

        Public Sub New(addedBuffers As IList(Of ITextBuffer), removedBuffers As IList(Of ITextBuffer))
            If addedBuffers Is Nothing Then
                Throw New ArgumentNullException("addedBuffers")
            End If

            If removedBuffers Is Nothing Then
                Throw New ArgumentNullException("removedBuffers")
            End If

            _addedBuffers = addedBuffers
            _removedBuffers = removedBuffers
        End Sub
    End Class
End Namespace
