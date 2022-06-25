Imports System.Collections.Generic

Namespace Microsoft.Nautilus.Text.Projection
    Public Interface IProjectionBuffer
        Inherits ITextBuffer, IPropertyOwner
        ReadOnly Property CurrentProjectionSnapshot As IProjectionSnapshot

        ReadOnly Property EditTransactionInProgress As Boolean

        ReadOnly Property SourceBuffers As IList(Of ITextBuffer)
        Event SourceSpansChanged As EventHandler(Of SpansChangedEventArgs)
        Event SourceBuffersChanged As EventHandler(Of BuffersChangedEventArgs)

        Sub InsertSpan(position As Integer, spanToInsert As ITextSpan)

        Sub InsertSpans(position As Integer, spansToInsert As IList(Of ITextSpan))

        Sub DeleteSpans(position As Integer, spansToDelete As Integer)

        Sub ReplaceSpans(position As Integer, spansToReplace As Integer, spanToInsert As ITextSpan)

        Sub ReplaceSpans(position As Integer, spansToReplace As Integer, spansToInsert As IList(Of ITextSpan))
    End Interface
End Namespace
