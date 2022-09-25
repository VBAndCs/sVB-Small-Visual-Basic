Imports System
Imports System.Collections.ObjectModel
Imports System.ComponentModel

Namespace Microsoft.SmallBasic.Documents
    Public Class DocumentTracker
        Private _openDocuments As New ObservableCollection(Of FileDocument)()

        Public ReadOnly Property OpenDocuments As ReadOnlyObservableCollection(Of FileDocument)
            Get
                Return New ReadOnlyObservableCollection(Of FileDocument)(_openDocuments)
            End Get
        End Property

        Public Shared Function GetDocumentProperty(Of T)(document As FileDocument) As T
            Dim value As Object = Nothing

            If document.PropertyStore.TryGetValue(GetType(T), value) Then
                Return value
            End If

            Return Nothing
        End Function

        Public Shared Sub SetDocumentProperty(document As FileDocument, value As Object)
            document.PropertyStore(value.GetType()) = value
        End Sub

        Public Sub TrackDocument(document As FileDocument)
            If document Is Nothing Then
                Throw New ArgumentNullException("document")
            End If

            _openDocuments.Add(document)
            AddHandler document.Closed,
                Sub()
                    _openDocuments.Remove(document)
                End Sub
        End Sub

    End Class
End Namespace
