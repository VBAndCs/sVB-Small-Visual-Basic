Imports System
Imports System.Collections.ObjectModel
Imports System.ComponentModel

Namespace Microsoft.SmallBasic.Documents
    Public Class DocumentTracker
        Private openDocumentsField As ObservableCollection(Of FileDocument) = New ObservableCollection(Of FileDocument)()

        Public ReadOnly Property OpenDocuments As ReadOnlyObservableCollection(Of FileDocument)
            Get
                Return New ReadOnlyObservableCollection(Of FileDocument)(openDocumentsField)
            End Get
        End Property

        Public Shared Function GetDocumentProperty(Of T)(ByVal document As FileDocument) As T
            Dim value As Object = Nothing

            If document.PropertyStore.TryGetValue(GetType(T), value) Then
                Return value
            End If

            Return Nothing
        End Function

        Public Shared Sub SetDocumentProperty(ByVal document As FileDocument, ByVal value As Object)
            document.PropertyStore(value.GetType()) = value
        End Sub

        Public Sub TrackDocument(ByVal document As FileDocument)
            If document Is Nothing Then
                Throw New ArgumentNullException("document")
            End If

            openDocumentsField.Add(document)
            AddHandler document.Closed,
                Sub()
                    openDocumentsField.Remove(document)
                End Sub
        End Sub

    End Class
End Namespace
