Imports System.Collections.Generic
Imports System.ComponentModel.Composition
Imports Microsoft.Nautilus.Core

Namespace Microsoft.Nautilus.Text.Document

    <Export(GetType(IBufferManagerFactory))>
    Public NotInheritable Class BufferManagerFactory
        Implements IBufferManagerFactory

        Private Property ContentSpecificBufferManagerFactories As IEnumerable(Of ImportInfo(Of ContentSpecificBufferManagerFactory, IContentTypeMetadata))

        Public Function CreateBufferManager(textBuffer As ITextBuffer) As IBufferManager Implements IBufferManagerFactory.CreateBufferManager
            Return CreateBufferManager(textBuffer, ContentSpecificBufferManagerFactories)
        End Function

        Friend Function CreateBufferManager(textBuffer As ITextBuffer, factories As IEnumerable(Of ImportInfo(Of ContentSpecificBufferManagerFactory, IContentTypeMetadata))) As IBufferManager
            Dim importInfo1 As ImportInfo(Of ContentSpecificBufferManagerFactory, IContentTypeMetadata) = MatchFactory(textBuffer.ContentType, factories)
            If importInfo1 IsNot Nothing Then
                Return importInfo1.GetBoundValue().CreateBufferManager(textBuffer)
            End If
            Return New VacuousBufferManager(textBuffer)
        End Function

        Friend Function MatchFactory(bufferContentType As String, factories As IEnumerable(Of ImportInfo(Of ContentSpecificBufferManagerFactory, IContentTypeMetadata))) As ImportInfo(Of ContentSpecificBufferManagerFactory, IContentTypeMetadata)
            For Each factory As ImportInfo(Of ContentSpecificBufferManagerFactory, IContentTypeMetadata) In factories
                For Each contentType1 As String In factory.Metadata.ContentTypes
                    If ContentTypeHelper.IsSame(bufferContentType, contentType1) Then
                        Return factory
                    End If
                Next
            Next

            If Not ContentTypeHelper.IsRootContentType(bufferContentType) Then
                Dim importInfo1 As ImportInfo(Of ContentSpecificBufferManagerFactory, IContentTypeMetadata) = MatchFactory(ContentTypeHelper.GetParentContentType(bufferContentType), factories)
                If importInfo1 IsNot Nothing Then
                    Return importInfo1
                End If
            End If

            Return Nothing
        End Function
    End Class
End Namespace
