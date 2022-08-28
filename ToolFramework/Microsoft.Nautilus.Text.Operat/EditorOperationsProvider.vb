Imports System.ComponentModel.Composition
Imports Microsoft.Nautilus.Text.Editor

Namespace Microsoft.Nautilus.Text.Operations

    <Export(GetType(IEditorOperationsProvider))>
    Public NotInheritable Class EditorOperationsProvider
        Implements IEditorOperationsProvider

        <Import>
        Public Property TextStructureNavigatorFactory As ITextStructureNavigatorFactory

        <Import>
        Public Property TextBufferUndoManagerProvider As ITextBufferUndoManagerProvider

        Public Function GetEditorOperations(textView As ITextView) As IEditorOperations Implements IEditorOperationsProvider.GetEditorOperations
            If textView Is Nothing Then
                Throw New ArgumentNullException("textView")
            End If

            Dim editorOperations As IEditorOperations = Nothing
            If Not textView.Properties.TryGetProperty(Of IEditorOperations)(GetType(EditorOperationsProvider), editorOperations) Then
                editorOperations = New EditorOperations(textView, _TextStructureNavigatorFactory)
                textView.Properties.AddProperty(GetType(EditorOperationsProvider), editorOperations)
            End If

            Return EditorOperations
        End Function
    End Class
End Namespace
