Imports System.Collections.Generic
Imports Microsoft.Nautilus.Core.Undo.DataExports
Imports System.Runtime.InteropServices

Namespace Microsoft.Nautilus.Core.Undo
    Public Interface IUndoHistoryRegistry
        ReadOnly Property Histories As IEnumerable(Of UndoHistory)

        ReadOnly Property GlobalHistory As UndoHistory

        ReadOnly Property UndoTransactionMarkers As IEnumerable(Of UndoTransactionMarker)

        Function RegisterHistory(context As Object) As UndoHistory

        Function RegisterHistory(context As Object, keepAlive As Boolean) As UndoHistory

        Function GetHistory(context As Object) As UndoHistory

        Function TryGetHistory(context As Object, <Out> ByRef history As UndoHistory) As Boolean

        Sub AttachHistory(context As Object, history As UndoHistory)

        Sub AttachHistory(context As Object, history As UndoHistory, keepAlive As Boolean)

        Sub DetachHistory(context As Object)

        Sub RemoveHistory(history As UndoHistory)

        Function GetTotalHistorySize() As Long

        Sub LimitHistoryDepth(depth As Integer)

        Sub LimitHistorySize(size As Long)

        Sub LimitTotalHistorySize(size As Long)

        Function GetUndoTransactionMarker(name As String) As UndoTransactionMarker

        Function CreateLinkedUndoTransaction(description As String) As LinkedUndoTransaction
    End Interface
End Namespace
