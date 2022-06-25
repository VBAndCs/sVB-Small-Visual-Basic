Imports System.ComponentModel.Composition
Imports Microsoft.Nautilus.Core.Undo

Namespace Microsoft.Nautilus.Text.Operations
    <Export(GetType(ITextBufferUndoManagerProvider))>
    Public NotInheritable Class TextBufferUndoManagerProvider
        Implements ITextBufferUndoManagerProvider

        <Import>
        Public Property UndoHistoryRegistry As IUndoHistoryRegistry

        Public Function GetTextBufferUndoManager(textBuffer As ITextBuffer) As ITextBufferUndoManager Implements ITextBufferUndoManagerProvider.GetTextBufferUndoManager
            If textBuffer Is Nothing Then
                Throw New ArgumentNullException("textBuffer")
            End If

            Dim [property] As TextBufferUndoManager = Nothing
            If Not textBuffer.Properties.TryGetProperty(GetType(ITextBufferUndoManager), [property]) Then
                [property] = New TextBufferUndoManager(textBuffer, _undoHistoryRegistry)
                textBuffer.Properties.AddProperty(GetType(ITextBufferUndoManager), [property])
            End If
            Return [property]
        End Function
    End Class
End Namespace
