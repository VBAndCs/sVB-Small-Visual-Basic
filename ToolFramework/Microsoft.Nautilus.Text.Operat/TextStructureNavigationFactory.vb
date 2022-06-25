Imports System.Collections.Generic
Imports System.ComponentModel.Composition
Imports Microsoft.Nautilus.Core

Namespace Microsoft.Nautilus.Text.Operations

    <Export(GetType(ITextStructureNavigatorFactory))>
    Public NotInheritable Class TextStructureNavigationFactory
        Implements ITextStructureNavigatorFactory

        <Import(GetType(TextStructureNavigatorCacheFactory))>
        Public Property TextStructureNavigatorProviders As IEnumerable(Of ImportInfo(Of Func(Of ITextBuffer, ITextStructureNavigator), IContentTypeMetadata))

        Public Function CreateTextStructureNavigator(textBuffer As ITextBuffer) As ITextStructureNavigator Implements ITextStructureNavigatorFactory.CreateTextStructureNavigator
            If textBuffer Is Nothing Then
                Throw New ArgumentNullException("textBuffer")
            End If

            Return CreateTextStructureNavigator(textBuffer, textBuffer.ContentType)
        End Function

        Public Function CreateTextStructureNavigator(textBuffer As ITextBuffer, contentType As String) As ITextStructureNavigator Implements ITextStructureNavigatorFactory.CreateTextStructureNavigator
            If textBuffer Is Nothing Then
                Throw New ArgumentNullException("textBuffer")
            End If

            If contentType Is Nothing Then
                Throw New ArgumentNullException("contentType")
            End If

            Dim navigatorProvider = GetNavigatorProvider(contentType)
            Dim textStructureNavigator As ITextStructureNavigator = Nothing

            If navigatorProvider IsNot Nothing Then
                textStructureNavigator = navigatorProvider.GetBoundValue()(textBuffer)
            End If

            If textStructureNavigator Is Nothing Then
                textStructureNavigator = New DefaultTextNavigator(textBuffer)
            End If

            Return textStructureNavigator
        End Function

        Private Function GetNavigatorProvider(contentType As String) As ImportInfo(Of Func(Of ITextBuffer, ITextStructureNavigator), IContentTypeMetadata)
            If String.IsNullOrEmpty(contentType) Then
                Return Nothing
            End If

            For Each textStructureNavigatorProvider As ImportInfo(Of Func(Of ITextBuffer, ITextStructureNavigator), IContentTypeMetadata) In _TextStructureNavigatorProviders
                For Each contentType2 As String In textStructureNavigatorProvider.Metadata.ContentTypes
                    If ContentTypeHelper.IsSame(contentType2, contentType) Then
                        Return textStructureNavigatorProvider
                    End If
                Next
            Next

            Return GetNavigatorProvider(ContentTypeHelper.GetParentContentType(contentType))
        End Function
    End Class
End Namespace
