Imports Microsoft.Nautilus.Core.Undo
Imports Microsoft.Nautilus.Text.Diagnostics
Imports Microsoft.Nautilus.Text.Editor

Namespace Microsoft.Nautilus.Text.Operations

    Public Class BeforeTextBufferChangeUndoPrimitive
        Inherits UndoPrimitive

        Private _textView As ITextView
        Private _oldCaretCharacterIndex As Integer
        Private _oldCaretPlacement As CaretPlacement
        Private _oldSelectionSpan As Span

        Public Overrides ReadOnly Property CanUndo As Boolean = True

        Public Overrides ReadOnly Property CanRedo As Boolean
            Get
                Return Not _canUndo
            End Get
        End Property

        Public Sub New(textView As ITextView, oldCaretCharacterIndex As Integer, oldCaretPlacement As CaretPlacement, oldSelectionSpan As Span)
            If textView Is Nothing Then
                Throw New ArgumentNullException("textView")
            End If

            If oldCaretCharacterIndex < 0 OrElse oldCaretCharacterIndex > textView.TextSnapshot.Length Then
                Throw New ArgumentOutOfRangeException("oldCaretCharacterIndex")
            End If

            _textView = textView
            _oldCaretCharacterIndex = oldCaretCharacterIndex
            _oldCaretPlacement = oldCaretPlacement
            _oldSelectionSpan = oldSelectionSpan
        End Sub

        Public Overrides Sub [Do]()
            EditorTrace.TraceBeforeTextBufferChangeRedoStart()

            If Not CanRedo Then
                Throw New InvalidOperationException("Cannot redo this change.")
            End If

            _canUndo = True
            EditorTrace.TraceBeforeTextBufferChangeRedoEnd()
        End Sub

        Public Overrides Sub Undo()
            EditorTrace.TraceBeforeTextBufferChangeUndoStart()
            If Not CanUndo Then
                Throw New InvalidOperationException("Cannot undo this change.")
            End If
            Try
                _textView.Caret.MoveTo(_oldCaretCharacterIndex, _oldCaretPlacement)
                _textView.Selection.ActiveSnapshotSpan = New SnapshotSpan(_textView.TextSnapshot, _oldSelectionSpan)
                _textView.Caret.EnsureVisible()
                _CanUndo = False
            Catch
            End Try
            EditorTrace.TraceBeforeTextBufferChangeUndoEnd()
        End Sub

    End Class
End Namespace
