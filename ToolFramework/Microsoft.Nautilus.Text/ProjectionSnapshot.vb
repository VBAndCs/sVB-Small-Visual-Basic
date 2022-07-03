Imports System.Collections.Generic
Imports System.IO
Imports System.Text
Imports Microsoft.Nautilus.Text.Projection
Imports Microsoft.Nautilus.Text.StringRebuilder

Namespace Microsoft.Nautilus.Text
    Friend Class ProjectionSnapshot
        Implements IProjectionSnapshot, ITextSnapshot

        Private projectionBuffer As IProjectionBuffer
        Private sourceSpans As List(Of SnapshotSpan)
        Private _sourceSnapshots As List(Of ITextSnapshot)
        Private lineBreakCounts As Integer()
        Private totalLength As Integer
        Private totalLineCount As Integer = 1

        Public ReadOnly Property SpanCount As Integer Implements IProjectionSnapshot.SpanCount
            Get
                Return sourceSpans.Count
            End Get
        End Property

        Public ReadOnly Property SourceSnapshots As IList(Of ITextSnapshot) Implements IProjectionSnapshot.SourceSnapshots
            Get
                Return _sourceSnapshots
            End Get
        End Property

        Public ReadOnly Property TextBuffer As ITextBuffer Implements ITextSnapshot.TextBuffer
            Get
                Return projectionBuffer
            End Get
        End Property

        Public ReadOnly Property Version As TextVersion Implements ITextSnapshot.Version

        Public ReadOnly Property Length As Integer Implements ITextSnapshot.Length
            Get
                Return totalLength
            End Get
        End Property

        Default Public ReadOnly Property Item(position As Integer) As Char Implements ITextSnapshot.Item
            Get
                Return MapToSourceSnapshot(position).GetChar()
            End Get
        End Property

        Public ReadOnly Property LineCount As Integer Implements ITextSnapshot.LineCount
            Get
                Return totalLineCount
            End Get
        End Property

        Public ReadOnly Iterator Property Lines As IEnumerable(Of ITextSnapshotLine) Implements ITextSnapshot.Lines
            Get
                For line As Integer = 0 To totalLineCount - 1
                    Yield GetLineFromLineNumber(line)
                Next
            End Get
        End Property

        Public Sub New(projectionBuffer As IProjectionBuffer, version As TextVersion, sourceSpans As List(Of SnapshotSpan))
            Me.projectionBuffer = projectionBuffer
            _Version = version
            Me.sourceSpans = sourceSpans
            lineBreakCounts = New Integer(sourceSpans.Count - 1) {}
            _sourceSnapshots = New List(Of ITextSnapshot)

            For i As Integer = 0 To sourceSpans.Count - 1
                Dim snapshotSpan1 As SnapshotSpan = sourceSpans(i)
                totalLength += snapshotSpan1.Length
                lineBreakCounts(i) = TextUtilities.ScanForLineCount(snapshotSpan1.GetText())
                totalLineCount += lineBreakCounts(i)
                Dim snapshot1 As ITextSnapshot = snapshotSpan1.Snapshot
                If Not _sourceSnapshots.Contains(snapshot1) Then
                    _sourceSnapshots.Add(snapshot1)
                End If
            Next
        End Sub

        Public Function GetMatchingSnapshot(textBuffer1 As ITextBuffer) As ITextSnapshot Implements IProjectionSnapshot.GetMatchingSnapshot
            If textBuffer1 Is Nothing Then
                Throw New ArgumentNullException("textBuffer")
            End If

            For Each sourceSnapshot As ITextSnapshot In _sourceSnapshots
                If sourceSnapshot.TextBuffer Is textBuffer1 Then
                    Return sourceSnapshot
                End If
            Next

            Return Nothing
        End Function

        Public Function GetSourceSpans() As IList(Of SnapshotSpan) Implements IProjectionSnapshot.GetSourceSpans
            Return sourceSpans
        End Function

        Public Function GetSourceSpans(startSpanIndex As Integer, count1 As Integer) As IList(Of SnapshotSpan) Implements IProjectionSnapshot.GetSourceSpans
            If startSpanIndex < 0 OrElse startSpanIndex > SpanCount Then
                Throw New ArgumentOutOfRangeException("startSpanIndex")
            End If

            If count1 < 0 OrElse startSpanIndex + count1 > SpanCount Then
                Throw New ArgumentOutOfRangeException("count")
            End If

            Dim list1 As New List(Of SnapshotSpan)(count1)
            For i As Integer = 0 To count1 - 1
                list1.Add(sourceSpans(startSpanIndex + i))
            Next
            Return list1
        End Function

        Friend Function GetSourceSpan(position As Integer) As SnapshotSpan
            Return sourceSpans(position)
        End Function

        Public Function MapToSourceSnapshots(span1 As Span) As IList(Of SnapshotSpan) Implements IProjectionSnapshot.MapToSourceSnapshots
            If span1.End > Length Then
                Throw New ArgumentOutOfRangeException("span")
            End If

            Dim spans As New List(Of SnapshotSpan)
            Dim lastPos As Integer = 0
            Dim flag As Boolean = True

            For Each sourceSpan In sourceSpans
                Dim pos = lastPos + sourceSpan.Length
                If flag Then
                    If span1.Start < pos OrElse (span1.Start = pos AndAlso span1.Start = Length) Then
                        flag = False
                        Dim start = sourceSpan.Start + span1.Start - lastPos
                        Dim [end] = If(start + span1.Length < sourceSpan.End, span1.Length, (sourceSpan.End - start))
                        spans.Add(New SnapshotSpan(sourceSpan.Snapshot, start, [end]))
                        If [end] = span1.Length Then Return spans
                    End If

                Else
                    If span1.End <= pos Then
                        spans.Add(New SnapshotSpan(sourceSpan.Snapshot, sourceSpan.Start, span1.End - lastPos))
                        Return spans
                    End If
                    spans.Add(sourceSpan)
                End If

                lastPos = pos
            Next

            Return spans
        End Function

        Public Function MapFromSourceSnapshot(sourceSpan As SnapshotSpan) As IList(Of Span) Implements IProjectionSnapshot.MapFromSourceSnapshot
            If Not _sourceSnapshots.Contains(sourceSpan.Snapshot) Then
                Throw New ArgumentException("The span does not belong to a source snapshot of the projection snapshot")
            End If

            Dim snapshot1 As ITextSnapshot = sourceSpan.Snapshot
            Dim list1 As New List(Of Span)
            Dim num As Integer = 0
            Dim num2 As Integer = 0

            For Each sourceSpan2 As SnapshotSpan In sourceSpans
                If sourceSpan2.Snapshot Is snapshot1 AndAlso sourceSpan2.OverlapsWith(sourceSpan) Then
                    Dim value1 = sourceSpan2.Overlap(sourceSpan).Value
                    Dim item As New Span(num2 + value1.Start - sourceSpan2.Start, value1.Length)
                    list1.Add(item)
                    num += item.Length
                    If num = sourceSpan.Length Then
                        Return list1
                    End If
                End If
                num2 += sourceSpan2.Length
            Next
            Return list1
        End Function

        Public Function MapToSourceSnapshot(position As Integer) As SnapshotPoint Implements IProjectionSnapshot.MapToSourceSnapshot
            If position < 0 OrElse position > Length Then
                Throw New ArgumentOutOfRangeException("position")
            End If

            Dim num As Integer = 0
            For Each sourceSpan As SnapshotSpan In sourceSpans
                Dim num2 As Integer = num + sourceSpan.Length
                If position < num2 OrElse (position = num2 AndAlso position = totalLength) Then
                    Return New SnapshotPoint(sourceSpan.Snapshot, sourceSpan.Start + position - num)
                End If
                num = num2
            Next

            Throw New ArgumentOutOfRangeException("position")
        End Function

        Public Function MapFromSourceSnapshot(sourcePoint As SnapshotPoint) As SnapshotPoint? Implements IProjectionSnapshot.MapFromSourceSnapshot
            If Not _sourceSnapshots.Contains(sourcePoint.Snapshot) Then
                Throw New ArgumentException("The point does not belong to a source snapshot of the projection snapshot")
            End If

            Dim num As Integer = 0
            For Each sourceSpan As SnapshotSpan In sourceSpans
                If sourceSpan.Snapshot Is sourcePoint.Snapshot AndAlso sourceSpan.Contains(sourcePoint.Position) Then
                    Return New SnapshotPoint(Me, num + sourcePoint.Position - sourceSpan.Start)
                End If
                num += sourceSpan.Length
            Next

            Return Nothing
        End Function

        Public Function GetText(span1 As Span) As String Implements ITextSnapshot.GetText
            If span1.End > totalLength Then
                Throw New ArgumentOutOfRangeException("span")
            End If

            Dim snapshotSpans = MapToSourceSnapshots(span1)
            If snapshotSpans.Count = 1 Then
                Return snapshotSpans(0).GetText()
            End If

            Dim sb As New StringBuilder
            For Each snapshotSpan In snapshotSpans
                sb.Append(snapshotSpan.GetText())
            Next
            Return sb.ToString()
        End Function

        Public Function GetText(startIndex As Integer, length1 As Integer) As String Implements ITextSnapshot.GetText
            Return GetText(New Span(startIndex, length1))
        End Function

        Public Function GetText(textSpan As ITextSpan) As String Implements ITextSnapshot.GetText
            If textSpan.TextBuffer IsNot projectionBuffer Then
                Throw New ArgumentException("textSpan belongs to the wrong TextBuffer")
            End If
            Return GetText(textSpan.GetSpan(Me))
        End Function

        Public Sub CopyTo(sourceIndex As Integer, destination As Char(), destinationIndex As Integer, count1 As Integer) Implements ITextSnapshot.CopyTo
            If destination Is Nothing Then
                Throw New ArgumentNullException("destination")
            End If

            If sourceIndex < 0 OrElse sourceIndex > totalLength Then
                Throw New ArgumentOutOfRangeException("sourceIndex")
            End If

            If count1 < 0 OrElse sourceIndex + count1 > totalLength OrElse destinationIndex + count1 > destination.Length Then
                Throw New ArgumentOutOfRangeException("count")
            End If

            If count1 <= 0 Then
                Return
            End If

            Dim list1 As IList(Of SnapshotSpan) = MapToSourceSnapshots(New Span(sourceIndex, count1))
            For Each item As SnapshotSpan In list1
                item.Snapshot.CopyTo(item.Start, destination, destinationIndex, item.Length)
                destinationIndex += item.Length
            Next
        End Sub

        Public Function ToCharArray(startIndex As Integer, length1 As Integer) As Char() Implements ITextSnapshot.ToCharArray
            If length1 < 0 OrElse length1 > totalLength Then
                Throw New ArgumentOutOfRangeException("length")
            End If

            Dim array As Char() = New Char(length1 - 1) {}
            CopyTo(startIndex, array, 0, length1)
            Return array
        End Function

        Public Function CreateTextPoint(position As Integer, trackingMode1 As TrackingMode) As ITextPoint Implements ITextSnapshot.CreateTextPoint
            Return New TextPoint(Me, position, trackingMode1)
        End Function

        Public Function CreateTextSpan(start1 As Integer, length1 As Integer, spanTrackingMode1 As SpanTrackingMode) As ITextSpan Implements ITextSnapshot.CreateTextSpan
            Return New TextSpan(Me, New Span(start1, length1), spanTrackingMode1)
        End Function

        Public Function CreateTextSpan(span1 As Span, spanTrackingMode1 As SpanTrackingMode) As ITextSpan Implements ITextSnapshot.CreateTextSpan
            Return New TextSpan(Me, span1, spanTrackingMode1)
        End Function

        Public Function GetLineFromLineNumber(lineNumber As Integer) As ITextSnapshotLine Implements ITextSnapshot.GetLineFromLineNumber
            If lineNumber >= totalLineCount Then
                Throw New ArgumentOutOfRangeException("lineNumber")
            End If
            Return GetLineFromPosition(PositionOfFirstCharacterInLine(lineNumber))
        End Function

        Public Function GetLineFromPosition(position As Integer) As ITextSnapshotLine Implements ITextSnapshot.GetLineFromPosition
            Dim num As Integer = 0
            Dim num2 As Integer = position

            While Math.Max(Threading.Interlocked.Decrement(num2), num2 + 1) > 0
                Select Case Me(num2)
                    Case vbLf(0), vbVerticalTab(0), vbFormFeed(0), ChrW(&H85)
                        num = num2 + 1
                        Exit While

                    Case vbCr(0)
                        If num2 + 1 < totalLength AndAlso Me(num2 + 1) = vbLf Then
                            Continue While
                        End If

                        num = num2 + 1
                        Exit While

                End Select

            End While

            Dim num3 As Integer = totalLength
            Dim lineBreakLength As Integer = 0
            For i As Integer = position To totalLength - 1
                Select Case Me(i)
                    Case vbVerticalTab(0), vbFormFeed(0), ChrW(&H85)
                        num3 = i
                        lineBreakLength = 1
                        Exit For

                    Case vbLf(0)
                        If i > 0 AndAlso Me(i - 1) = vbCr Then
                            num3 = i - 1
                            lineBreakLength = 2
                        Else
                            num3 = i
                            lineBreakLength = 1
                        End If
                        Exit For

                    Case vbCr(0)
                        num3 = i
                        lineBreakLength = (If((i >= totalLength - 1 OrElse Me(i + 1) <> vbLf), 1, 2))
                        Exit For

                End Select

            Next

            Return New TextSnapshotLine(Me, New LineSpan(GetLineNumberFromPosition(position), New Span(num, num3 - num), lineBreakLength))
        End Function

        Public Function GetLineNumberFromPosition(position As Integer) As Integer Implements ITextSnapshot.GetLineNumberFromPosition
            If position < 0 OrElse position > totalLength Then
                Throw New ArgumentOutOfRangeException("position")
            End If

            Dim num As Integer = 0
            Dim num2 As Integer = 0
            For i As Integer = 0 To sourceSpans.Count - 1
                Dim snapshotSpan1 As SnapshotSpan = sourceSpans(i)
                If position < num + snapshotSpan1.Length Then
                    Dim num3 As Integer = position - num + 1
                    If num3 = 1 Then
                        Return num2
                    End If
                    Dim text As String = snapshotSpan1.Snapshot.GetText(snapshotSpan1.Start, num3)
                    Dim num4 As Integer = TextUtilities.ScanForLineCount(text)
                    Dim c As Char = Me(position)
                    If c = vbCr OrElse c = vbLf OrElse c = vbVerticalTab OrElse c = vbFormFeed OrElse c = ChrW(&H85) Then
                        num4 -= 1
                    End If
                    Return num2 + num4
                End If
                num += snapshotSpan1.Length
                num2 += lineBreakCounts(i)
            Next
            Return num2
        End Function

        Private Function PositionOfFirstCharacterInLine(lineNumber As Integer) As Integer
            If lineNumber = 0 Then
                Return 0
            End If

            Dim flag As Boolean = False
            Dim num As Integer = 0
            For i As Integer = 0 To sourceSpans.Count - 1
                Dim snapshotSpan1 As SnapshotSpan = sourceSpans(i)
                If snapshotSpan1.Length <= 0 Then
                    Continue For
                End If
                If flag Then
                    Return MapFromSourceSnapshot(New SnapshotPoint(snapshotSpan1.Snapshot, snapshotSpan1.Start)).Value
                End If
                If lineNumber = num + lineBreakCounts(i) Then
                    Dim c As Char = snapshotSpan1.Snapshot(snapshotSpan1.End - 1)
                    If c = vbLf OrElse c = vbVerticalTab OrElse c = vbFormFeed OrElse c = ChrW(&H85) Then
                        flag = True
                        Continue For
                    End If
                End If
                If lineNumber < num + lineBreakCounts(i) Then
                    Dim lineNumberFromPosition As Integer = snapshotSpan1.Snapshot.GetLineNumberFromPosition(snapshotSpan1.Start)
                    Dim start1 As Integer = snapshotSpan1.Snapshot.GetLineFromLineNumber(lineNumberFromPosition + lineNumber - num).Start
                    start1 = Math.Max(start1, snapshotSpan1.Start)
                    Return MapFromSourceSnapshot(New SnapshotPoint(snapshotSpan1.Snapshot, start1)).Value
                End If
                num += lineBreakCounts(i)
            Next
            Return Length
        End Function

        Public Sub Write(writer As TextWriter) Implements ITextSnapshot.Write
            If writer Is Nothing Then
                Throw New ArgumentNullException("writer")
            End If
            UncheckedWrite(writer, New Span(0, totalLength))
        End Sub

        Public Sub Write(writer As TextWriter, span1 As Span) Implements ITextSnapshot.Write
            If writer Is Nothing Then
                Throw New ArgumentNullException("writer")
            End If
            If span1.End > totalLength Then
                Throw New ArgumentOutOfRangeException("span")
            End If
            UncheckedWrite(writer, span1)
        End Sub

        Private Sub UncheckedWrite(writer As TextWriter, span1 As Span)
            Dim list1 As IList(Of SnapshotSpan) = MapToSourceSnapshots(span1)
            For Each item As SnapshotSpan In list1
                item.Snapshot.Write(writer, item.Span)
            Next
        End Sub
    End Class
End Namespace
