Imports System.Collections
Imports System.Collections.Generic

Namespace Microsoft.Nautilus.Text.Editor
    Public Interface ITextSelection
        Inherits IEnumerable(Of ITextSpan), IEnumerable
        ReadOnly Property Count As Integer

        ReadOnly Property IsEmpty As Boolean
        Default ReadOnly Property Item(index As Integer) As ITextSpan

        Property ActiveSpan As ITextSpan

        Property ActiveSnapshotSpan As SnapshotSpan

        Property IsActiveSpanReversed As Boolean
        Event SelectionChanged As EventHandler

        Sub [Select](span As ITextSpan)

        Sub Clear()
    End Interface
End Namespace
