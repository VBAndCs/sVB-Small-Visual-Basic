Imports System.Collections.Generic
Imports System.IO
Imports Microsoft.Nautilus.Text.StringRebuilder

Namespace Microsoft.Nautilus.Text
    Friend Class TextSnapshot
        Implements ITextSnapshot

        Private content As IStringRebuilder

        Public ReadOnly Property TextBuffer As ITextBuffer Implements ITextSnapshot.TextBuffer

        Public ReadOnly Property Version As TextVersion Implements ITextSnapshot.Version

        Public ReadOnly Property Length As Integer Implements ITextSnapshot.Length
            Get
                Return content.Length
            End Get
        End Property

        Public ReadOnly Property LineCount As Integer Implements ITextSnapshot.LineCount
            Get
                Return content.LineBreakCount + 1
            End Get
        End Property

        Default Public ReadOnly Property Item(position As Integer) As Char Implements ITextSnapshot.Item
            Get
                Return content(position)
            End Get
        End Property

        Public ReadOnly Iterator Property Lines As IEnumerable(Of ITextSnapshotLine) Implements ITextSnapshot.Lines
            Get
                For line As Integer = 0 To LineCount - 1
                    Yield GetLineFromLineNumber(line)
                Next
            End Get
        End Property

        Public Sub New(textBuffer As ITextBuffer, version As TextVersion, content As IStringRebuilder)
            _TextBuffer = textBuffer
            _Version = version
            Me.content = content
        End Sub

        Public Function GetText(span1 As Span) As String Implements ITextSnapshot.GetText
            Return content.GetText(span1)
        End Function

        Public Function GetText(startIndex As Integer, length1 As Integer) As String Implements ITextSnapshot.GetText
            Return GetText(New Span(startIndex, length1))
        End Function

        Public Function GetText(textSpan As ITextSpan) As String Implements ITextSnapshot.GetText
            If textSpan Is Nothing Then
                Throw New ArgumentNullException("textSpan")
            End If
            Return GetText(textSpan.GetSpan(Me))
        End Function

        Public Sub CopyTo(sourceIndex As Integer, destination As Char(), destinationIndex As Integer, count As Integer) Implements ITextSnapshot.CopyTo
            content.CopyTo(sourceIndex, destination, destinationIndex, count)
        End Sub

        Public Function ToCharArray(startIndex As Integer, length1 As Integer) As Char() Implements ITextSnapshot.ToCharArray
            Return content.ToCharArray(startIndex, length1)
        End Function

        Public Function CreateTextPoint(position As Integer, trackingMode1 As TrackingMode) As ITextPoint Implements ITextSnapshot.CreateTextPoint
            Return New TextPoint(Me, position, trackingMode1)
        End Function

        Public Function CreateTextSpan(start As Integer, length1 As Integer, spanTrackingMode1 As SpanTrackingMode) As ITextSpan Implements ITextSnapshot.CreateTextSpan
            Return New TextSpan(Me, start, length1, spanTrackingMode1)
        End Function

        Public Function CreateTextSpan(span1 As Span, spanTrackingMode1 As SpanTrackingMode) As ITextSpan Implements ITextSnapshot.CreateTextSpan
            Return New TextSpan(Me, span1, spanTrackingMode1)
        End Function

        Public Function GetLineFromLineNumber(lineNumber As Integer) As ITextSnapshotLine Implements ITextSnapshot.GetLineFromLineNumber
            Dim line = content.GetLineFromLineNumber(lineNumber)
            Return New TextSnapshotLine(Me, line)
        End Function

        Public Function GetLineFromPosition(position As Integer) As ITextSnapshotLine Implements ITextSnapshot.GetLineFromPosition
            Dim lineNumberFromPosition As Integer = content.GetLineNumberFromPosition(position)
            Return GetLineFromLineNumber(lineNumberFromPosition)
        End Function

        Public Function GetLineNumberFromPosition(position As Integer) As Integer Implements ITextSnapshot.GetLineNumberFromPosition
            Return content.GetLineNumberFromPosition(position)
        End Function

        Public Sub Write(writer As TextWriter) Implements ITextSnapshot.Write
            content.Write(writer, New Span(0, content.Length))
        End Sub

        Public Sub Write(writer As TextWriter, span1 As Span) Implements ITextSnapshot.Write
            content.Write(writer, span1)
        End Sub
    End Class
End Namespace
