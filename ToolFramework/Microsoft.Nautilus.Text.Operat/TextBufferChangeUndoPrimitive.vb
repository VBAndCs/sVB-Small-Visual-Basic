Imports Microsoft.Nautilus.Core.Undo
Imports Microsoft.Nautilus.Text.Diagnostics

Namespace Microsoft.Nautilus.Text.Operations
    Public Class TextBufferChangeUndoPrimitive
        Inherits UndoPrimitive

        Private _textBuffer As ITextBuffer
        Private _textChange As ITextChange

        Public Overrides ReadOnly Property CanUndo As Boolean = True

        Public Overrides ReadOnly Property CanRedo As Boolean
            Get
                Return Not _canUndo
            End Get
        End Property

        Public Sub New(textBuffer As ITextBuffer, textChange As ITextChange)
            If textBuffer Is Nothing Then
                Throw New ArgumentNullException("textBuffer")
            End If

            If textChange Is Nothing Then
                Throw New ArgumentNullException("textChange")
            End If

            _textBuffer = textBuffer
            _textChange = textChange
        End Sub

        Public Overrides Sub [Do]()
            EditorTrace.TraceTextBufferChangeRedoStart()

            If Not CanRedo Then
                Throw New InvalidOperationException("Cannot redo this change.")
            End If

            _textBuffer.Replace(New Span(_textChange.Position, _textChange.OldLength), _textChange.NewText)
            _canUndo = True

            EditorTrace.TraceTextBufferChangeRedoEnd()
        End Sub

        Public Overrides Sub Undo()
            EditorTrace.TraceTextBufferChangeUndoStart()

            If Not CanUndo Then
                Throw New InvalidOperationException("Cannot undo this change.")
            End If

            Dim textSnapshot As ITextSnapshot = _textBuffer.Replace(New Span(_textChange.Position, _textChange.NewLength), _textChange.OldText)
            _canUndo = False

            EditorTrace.TraceTextBufferChangeUndoEnd()
        End Sub
    End Class
End Namespace
