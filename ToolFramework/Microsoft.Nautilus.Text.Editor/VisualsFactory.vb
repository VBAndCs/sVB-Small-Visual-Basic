Imports System.Collections.Generic
Imports System.Windows.Media.TextFormatting
Imports Microsoft.Nautilus.Text.Classification
Imports Microsoft.Nautilus.Text.Diagnostics

Namespace Microsoft.Nautilus.Text.Editor
    Friend Class VisualsFactory
        Private _classifier As IClassifier
        Private _classificationFormatMap As IClassificationFormatMap
        Private _textView As AvalonTextView
        Private _classificationTypeRegistry As IClassificationTypeRegistry
        Private _paragraphProperties As New TextFormattingParagraphProperties
        Private Shared _textFormatter As TextFormatter

        Public Sub New(textView As AvalonTextView, classifier As IClassifier, classificationFormatMap As IClassificationFormatMap, classificationTypeRegistry As IClassificationTypeRegistry)
            _textView = textView
            _classifier = classifier
            _classificationFormatMap = classificationFormatMap
            _classificationTypeRegistry = classificationTypeRegistry
        End Sub

        Public Function CreateLineVisuals(
                               line As ITextSnapshotLine,
                               position As Integer,
                               horizontalOffset As Double,
                               verticalOffset As Double,
                               maxIndent As Double,
                               maxWidth As Double,
                               spaceNegotiations As IList(Of SpaceNegotiation)
                    ) As IList(Of TextLineVisual)

            EditorTrace.TraceTextRenderStart()
            Dim lineEnd = line.EndIncludingLineBreak
            Dim positions As New List(Of Integer)

            Dim snapshot As New SnapshotSpan(_textView.TextSnapshot, line.ExtentIncludingLineBreak)
            Dim classificationSpans = If(lineEnd <= position,
                        New List(Of ClassificationSpan)(0),
                        _classifier.GetClassificationSpans(snapshot)
                    )

            Dim formattingSource As New TextFormattingSource(
                    _textView.TextSnapshot,
                    line.ExtentIncludingLineBreak,
                    _classificationFormatMap,
                    classificationSpans,
                    spaceNegotiations,
                    _classificationTypeRegistry)

            For Each pos In formattingSource.VirtualCharacterPositions
                positions.Add(pos + position)
            Next

            positions.Sort()
            If _textFormatter Is Nothing Then
                _textFormatter = TextFormatter.Create()
            End If

            Dim curPos As Integer = position
            Dim index As Integer = 0
            Dim lineVisuals As New List(Of TextLineVisual)

            Do
                Dim lines As New List(Of TextLine)
                Dim lineLength = 0
                Dim nextPos = curPos
                Dim nextIndex = index
                Dim previousLineBreak As TextLineBreak = Nothing

                Do
                    Dim formatedLine = _textFormatter.FormatLine(formattingSource, nextIndex, maxWidth, _paragraphProperties, previousLineBreak)
                    previousLineBreak = formatedLine.GetTextLineBreak()
                    lines.Add(New TextLine(formatedLine))
                    Dim endPos = formatedLine.Length
                    nextIndex += endPos

                    For Each pos In positions
                        If pos >= curPos Then
                            If pos >= curPos + endPos Then Exit For
                            endPos -= 1
                        End If
                    Next

                    nextPos += endPos
                    If nextPos >= lineEnd Then
                        lineLength = line.LineBreakLength
                        nextPos = lineEnd
                        Exit Do
                    End If
                Loop While maxWidth = 0.0

                Dim lineVisual As New TextLineVisual(
                        _textView,
                        lines,
                        index,
                        New Span(curPos, nextPos - curPos),
                        lineLength,
                        positions,
                        horizontalOffset,
                        verticalOffset
                 )
                lineVisuals.Add(lineVisual)

                If nextPos >= lineEnd Then Exit Do

                verticalOffset += lineVisual.Height
                If curPos = line.Start AndAlso maxIndent > 0.0 Then
                    Dim indent = lineVisual.GetIndentation()
                    If indent > maxIndent Then indent = maxIndent
                    horizontalOffset += indent
                    maxWidth -= indent
                    maxIndent = 0.0
                End If

                curPos = nextPos
                index = nextIndex
            Loop

            EditorTrace.TraceTextRenderEnd()
            Return lineVisuals
        End Function
    End Class
End Namespace
