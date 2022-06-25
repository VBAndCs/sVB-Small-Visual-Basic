Imports Microsoft.Nautilus.Core.Undo
Imports Microsoft.Nautilus.Text.Diagnostics

Namespace Microsoft.Nautilus.Text.Operations
    Public Class ReadOnlyRegionChangedUndoPrimitive
        Inherits UndoPrimitive

        Private _readOnlyRegion As IReadOnlyRegionHandle
        Private _change As ReadOnlyRegionChange

        Public Overrides ReadOnly Property CanUndo As Boolean = True

        Public Overrides ReadOnly Property CanRedo As Boolean
            Get
                Return Not _canUndo
            End Get
        End Property

        Public Sub New(readOnlyRegion As IReadOnlyRegionHandle, change As ReadOnlyRegionChange)
            If readOnlyRegion Is Nothing Then
                Throw New ArgumentNullException("readOnlyRegion")
            End If

            _readOnlyRegion = readOnlyRegion
            _change = change
        End Sub

        Private Sub RecreateReadOnlyRegion()
            Dim textBuffer1 As ITextBuffer = _readOnlyRegion.Span.TextBuffer
            Dim currentSnapshot1 As ITextSnapshot = textBuffer1.CurrentSnapshot
            Dim span1 As SnapshotSpan = _readOnlyRegion.Span.GetSpan(currentSnapshot1)
            _readOnlyRegion = textBuffer1.ReadOnlyRegionManager.CreateReadOnlyRegion(span1)
        End Sub

        Private Sub RemoveReadOnlyRegion()
            _readOnlyRegion.Remove()
        End Sub

        Public Overrides Sub [Do]()
            EditorTrace.TraceTextBufferChangeRedoStart()

            If Not CanRedo Then
                Throw New InvalidOperationException("Cannot redo this change.")
            End If

            If _change = ReadOnlyRegionChange.Created Then
                RecreateReadOnlyRegion()
            ElseIf _change = ReadOnlyRegionChange.Removed Then
                RemoveReadOnlyRegion()
            End If

            _CanUndo = True

            EditorTrace.TraceTextBufferChangeRedoEnd()
        End Sub

        Public Overrides Sub Undo()
            EditorTrace.TraceTextBufferChangeUndoStart()

            If Not CanUndo Then
                Throw New InvalidOperationException("Cannot undo this change.")
            End If

            If _change = ReadOnlyRegionChange.Created Then
                RemoveReadOnlyRegion()
            ElseIf _change = ReadOnlyRegionChange.Removed Then
                RecreateReadOnlyRegion()
            End If

            _CanUndo = False

            EditorTrace.TraceTextBufferChangeUndoEnd()
        End Sub
    End Class
End Namespace
