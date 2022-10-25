Imports System.IO

Namespace Microsoft.Nautilus.Text
    Public Interface ITextSnapshot
        ReadOnly Property TextBuffer As ITextBuffer

        ReadOnly Property Version As TextVersion

        ReadOnly Property Length As Integer

        ReadOnly Property LineCount As Integer
        Default ReadOnly Property Item(position As Integer) As Char

        ReadOnly Property Lines As IEnumerable(Of ITextSnapshotLine)

        Function GetText(span As Span) As String

        Function GetText(startIndex As Integer, length As Integer) As String

        Function GetText(textSpan As ITextSpan) As String

        Function ToCharArray(startIndex As Integer, length As Integer) As Char()

        Sub CopyTo(sourceIndex As Integer, destination As Char(), destinationIndex As Integer, count As Integer)

        Function CreateTextPoint(position As Integer, trackingMode As TrackingMode) As ITextPoint

        Function CreateTextSpan(span As Span, spanTrackingMode As SpanTrackingMode) As ITextSpan

        Function CreateTextSpan(start As Integer, length As Integer, spanTrackingMode As SpanTrackingMode) As ITextSpan

        Function GetLineFromLineNumber(lineNumber As Integer) As ITextSnapshotLine

        Function GetLineFromPosition(position As Integer) As ITextSnapshotLine

        Function GetLineNumberFromPosition(position As Integer) As Integer

        Sub Write(writer As TextWriter, span As Span)

        Sub Write(writer As TextWriter)
    End Interface
End Namespace
