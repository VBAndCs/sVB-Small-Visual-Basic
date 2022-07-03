Imports Microsoft.Nautilus.Core.Undo

Namespace Microsoft.Nautilus.Text.Operations
    Friend Class TextBufferUndoManager
        Implements ITextBufferUndoManager, IDisposable

        Private _undoHistoryRegistry As IUndoHistoryRegistry
        Private _undoHistory As UndoHistory

        Public ReadOnly Property TextBuffer As ITextBuffer Implements ITextBufferUndoManager.TextBuffer

        Public ReadOnly Property TextBufferUndoHistory As UndoHistory Implements ITextBufferUndoManager.TextBufferUndoHistory
            Get
                _undoHistory = _undoHistoryRegistry.RegisterHistory(_textBuffer)
                Return _undoHistory
            End Get
        End Property

        Public Sub New(textBuffer As ITextBuffer, undoHistoryRegistry As IUndoHistoryRegistry)
            If textBuffer Is Nothing Then
                Throw New ArgumentNullException("textBuffer")
            End If

            If undoHistoryRegistry Is Nothing Then
                Throw New ArgumentNullException("undoHistoryRegistry")
            End If

            _TextBuffer = textBuffer
            _undoHistoryRegistry = undoHistoryRegistry
            _undoHistory = _undoHistoryRegistry.RegisterHistory(_TextBuffer)

            AddHandler _TextBuffer.Changed, AddressOf TextBufferChanged
            AddHandler _TextBuffer.ReadOnlyRegionManager.ReadOnlyRegionChanged, AddressOf ReadOnlyRegionChanged
        End Sub

        Private Sub TextBufferChanged(sender As Object, e As TextChangedEventArgs)
            If _undoHistory.State <> 0 Then Return

            Using undoTransaction = _undoHistory.CreateTransaction("Text Buffer Change")
                For Each change In e.Changes
                    Dim undo As New TextBufferChangeUndoPrimitive(_TextBuffer, change)
                    undoTransaction.AddUndo(undo)
                Next
                undoTransaction.Complete()
            End Using
        End Sub

        Private Sub ReadOnlyRegionChanged(sender As Object, e As ReadOnlyRegionChangedEventArgs)
            If _undoHistory.State = UndoHistoryState.Idle Then
                Using undoTransaction = _undoHistory.CreateTransaction(If((e.Change = ReadOnlyRegionChange.Created), "Create Read Only Region", "Remove Read Only Region"))
                    Dim undo As New ReadOnlyRegionChangedUndoPrimitive(e.ReadOnlyRegion, e.Change)
                    undoTransaction.AddUndo(undo)
                    undoTransaction.Complete()
                End Using
            End If
        End Sub

        Public Sub UnregisterUndoHistory() Implements ITextBufferUndoManager.UnregisterUndoHistory
            _undoHistoryRegistry.RemoveHistory(_undoHistory)
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            RemoveHandler _textBuffer.Changed, AddressOf TextBufferChanged
            RemoveHandler _textBuffer.ReadOnlyRegionManager.ReadOnlyRegionChanged, AddressOf ReadOnlyRegionChanged
        End Sub
    End Class
End Namespace
