Imports Microsoft.Nautilus.Core.Undo
Imports Microsoft.Nautilus.Text.Editor

Namespace Microsoft.Nautilus.Text.Operations

    Public Class AfterTextBufferChangeUndoPrimitive
        Inherits UndoPrimitive

        Private _textView As ITextView
        Private _newCaretCharacterIndex As Integer
        Private _newCaretPlacement As CaretPlacement
        Private _newSelectionSpan As Span

        Public Overrides ReadOnly Property CanUndo As Boolean = True

        Public Overrides ReadOnly Property CanRedo As Boolean
            Get
                Return Not _canUndo
            End Get
        End Property

        Public Sub New(
                      textView As ITextView,
                      newCaretCharacterIndex As Integer,
                      newCaretPlacement As CaretPlacement,
                      newSelectionSpan As Span)

            If textView Is Nothing Then
                Throw New ArgumentNullException("textView")
            End If

            If newCaretCharacterIndex < 0 OrElse newCaretCharacterIndex > textView.TextSnapshot.Length Then
                Throw New ArgumentOutOfRangeException("newCaretCharacterIndex")
            End If

            _textView = textView
            _newCaretCharacterIndex = newCaretCharacterIndex
            _newCaretPlacement = newCaretPlacement
            _newSelectionSpan = newSelectionSpan
        End Sub

        Public Overrides Sub [Do]()
            If Not CanRedo Then
                Throw New InvalidOperationException("Cannot redo this change.")
            End If

            _textView.Caret.MoveTo(_newCaretCharacterIndex, _newCaretPlacement)
            _textView.Selection.ActiveSnapshotSpan = New SnapshotSpan(_textView.TextSnapshot, _newSelectionSpan)
            _textView.Caret.EnsureVisible()
            _canUndo = True
        End Sub

        Public Overrides Sub Undo()
            If Not _CanUndo Then
                Throw New InvalidOperationException("Cannot undo this change.")
            End If

            _CanUndo = False
        End Sub
    End Class
End Namespace
