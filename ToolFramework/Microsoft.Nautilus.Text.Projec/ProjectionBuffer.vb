Imports System.Collections.Generic
Imports System.Text
Imports System.Runtime.InteropServices

Namespace Microsoft.Nautilus.Text.Projection.Implementation
    Friend Class ProjectionBuffer
        Inherits BaseBuffer
        Implements IProjectionBuffer, ITextBuffer, IPropertyOwner

        Private Class SourceBufferSet
            Private Class BufferTracker
                Public _buffer As ITextBuffer
                Public _spanCount As Integer

                Public Sub New(buffer As ITextBuffer)
                    _buffer = buffer
                    _spanCount = 1
                End Sub
            End Class

            Private _inTransaction As Boolean
            Private _sourceBufferTrackers As New List(Of BufferTracker)
            Private _addedBuffers As List(Of ITextBuffer)
            Private _removedBuffers As List(Of ITextBuffer)

            Public ReadOnly Property SourceBuffers As IList(Of ITextBuffer)
                Get
                    Dim list1 As New List(Of ITextBuffer)
                    For Each sourceBufferTracker As BufferTracker In _sourceBufferTrackers
                        If Not list1.Contains(sourceBufferTracker._buffer) Then
                            list1.Add(sourceBufferTracker._buffer)
                        End If
                    Next

                    Return list1
                End Get
            End Property

            Private Function Find(buffer As ITextBuffer) As Integer
                For i As Integer = 0 To _sourceBufferTrackers.Count - 1
                    If buffer Is _sourceBufferTrackers(i)._buffer Then
                        Return i
                    End If
                Next

                Return -1
            End Function

            Public Function Contains(buffer As ITextBuffer) As Boolean
                Return Find(buffer) >= 0
            End Function

            Public Sub StartTransaction()
                _addedBuffers = New List(Of ITextBuffer)
                _removedBuffers = New List(Of ITextBuffer)
                _inTransaction = True
            End Sub

            Public Sub FinishTransaction(<Out> ByRef addedBuffers As List(Of ITextBuffer), <Out> ByRef removedBuffers As List(Of ITextBuffer))
                addedBuffers = _addedBuffers
                removedBuffers = _removedBuffers
                _addedBuffers = Nothing
                _removedBuffers = Nothing
                _inTransaction = False
            End Sub

            Public Sub AddSpan(span As ITextSpan)
                Dim num As Integer = Find(span.TextBuffer)
                If num < 0 Then
                    _sourceBufferTrackers.Add(New BufferTracker(span.TextBuffer))
                    _addedBuffers.Add(span.TextBuffer)
                Else
                    _sourceBufferTrackers(num)._spanCount += 1
                End If
            End Sub

            Public Sub RemoveSpan(span As ITextSpan)
                Dim index As Integer = Find(span.TextBuffer)
                If Threading.Interlocked.Decrement(_sourceBufferTrackers(index)._spanCount) = 0 Then
                    _sourceBufferTrackers.RemoveAt(index)
                    _removedBuffers.Add(span.TextBuffer)
                End If
            End Sub
        End Class

        Private sourceSpans As New List(Of ITextSpan)
        Private _sourceBufferSet As New SourceBufferSet
        Private _currentProjectionSnapshot As ProjectionSnapshot
        Private pendingSourceChanges As New Dictionary(Of ITextBuffer, List(Of ITextChange))

        Public ReadOnly Property EditTransactionInProgress As Boolean Implements IProjectionBuffer.EditTransactionInProgress

        Public ReadOnly Property CurrentProjectionSnapshot As IProjectionSnapshot Implements IProjectionBuffer.CurrentProjectionSnapshot
            Get
                Return _currentProjectionSnapshot
            End Get
        End Property

        Public ReadOnly Property SourceBuffers As IList(Of ITextBuffer) Implements IProjectionBuffer.SourceBuffers
            Get
                Return _sourceBufferSet.SourceBuffers
            End Get
        End Property

        Public Event SourceBuffersChanged As EventHandler(Of BuffersChangedEventArgs) Implements IProjectionBuffer.SourceBuffersChanged

        Public Event SourceSpansChanged As EventHandler(Of SpansChangedEventArgs) Implements IProjectionBuffer.SourceSpansChanged

        Public Sub New(resolver As IProjectionEditResolver, contentType As String, initialSourceSpans As IEnumerable(Of ITextSpan))
            MyBase.New(contentType)
            If initialSourceSpans Is Nothing Then
                Throw New ArgumentNullException("sourceSpans")
            End If

            CheckInsertionLegality(initialSourceSpans)
            Dim num As Integer = 0
            _sourceBufferSet.StartTransaction()
            Dim list1 As New List(Of SnapshotSpan)

            For Each initialSourceSpan As ITextSpan In initialSourceSpans
                AddSpan(Math.Min(Threading.Interlocked.Increment(num), num - 1), initialSourceSpan)
                list1.Add(New SnapshotSpan(initialSourceSpan.TextBuffer.CurrentSnapshot, initialSourceSpan.GetSpan(initialSourceSpan.TextBuffer.CurrentSnapshot)))
            Next

            Dim addedBuffers As New List(Of ITextBuffer)
            _sourceBufferSet.FinishTransaction(addedBuffers, New List(Of ITextBuffer))

            For Each item As ITextBuffer In addedBuffers
                AddHandler item.Changed, AddressOf OnSourceTextChanged
            Next

            _currentProjectionSnapshot = New ProjectionSnapshot(Me, MyBase.CurrentVersion, list1)
            MyBase.CurrentSnapshot = CurrentProjectionSnapshot
        End Sub

        Public Sub InsertSpan(position As Integer, spanToInsert As ITextSpan) Implements IProjectionBuffer.InsertSpan
            Dim list1 As New List(Of ITextSpan)(1)
            list1.Add(spanToInsert)
            InsertSpans(position, list1)
        End Sub

        Public Sub InsertSpans(position As Integer, spansToInsert As IList(Of ITextSpan)) Implements IProjectionBuffer.InsertSpans
            If position < 0 OrElse position > sourceSpans.Count Then
                Throw New ArgumentOutOfRangeException("position")
            End If

            If spansToInsert Is Nothing Then
                Throw New ArgumentNullException("spansToInsert")
            End If

            If spansToInsert.Count > 0 Then
                PerformSpanReplacement(position, 0, spansToInsert)
            End If
        End Sub

        Public Sub DeleteSpans(position As Integer, spansToDelete As Integer) Implements IProjectionBuffer.DeleteSpans
            If position < 0 OrElse position > sourceSpans.Count Then
                Throw New ArgumentOutOfRangeException("position")
            End If

            If spansToDelete < 0 OrElse position + spansToDelete > sourceSpans.Count Then
                Throw New ArgumentOutOfRangeException("spansToDelete")
            End If

            PerformSpanReplacement(position, spansToDelete, New List(Of ITextSpan))
        End Sub

        Public Sub ReplaceSpans(position As Integer, spansToReplace As Integer, spanToInsert As ITextSpan) Implements IProjectionBuffer.ReplaceSpans
            If position < 0 OrElse position > sourceSpans.Count Then
                Throw New ArgumentOutOfRangeException("position")
            End If

            If spansToReplace < 0 OrElse position + spansToReplace > sourceSpans.Count Then
                Throw New ArgumentOutOfRangeException("spansToReplace")
            End If

            If spanToInsert Is Nothing Then
                Throw New ArgumentNullException("spansToInsert")
            End If

            Dim list1 As New List(Of ITextSpan)(1)
            list1.Add(spanToInsert)
            PerformSpanReplacement(position, spansToReplace, list1)
        End Sub

        Public Sub ReplaceSpans(position As Integer, spansToReplace As Integer, spansToInsert As IList(Of ITextSpan)) Implements IProjectionBuffer.ReplaceSpans
            If position < 0 OrElse position > sourceSpans.Count Then
                Throw New ArgumentOutOfRangeException("position")
            End If

            If spansToReplace < 0 OrElse position + spansToReplace > sourceSpans.Count Then
                Throw New ArgumentOutOfRangeException("spansToReplace")
            End If

            If spansToInsert Is Nothing Then
                Throw New ArgumentNullException("spansToInsert")
            End If
            PerformSpanReplacement(position, spansToReplace, spansToInsert)
        End Sub

        Private Sub PerformSpanReplacement(position As Integer, spansToDelete As Integer, spansToInsert As IList(Of ITextSpan))
            If spansToInsert.Count > 0 Then
                CheckInsertionLegality(spansToInsert)
            End If

            Dim list1 As New List(Of ITextSpan)
            _sourceBufferSet.StartTransaction()
            Dim stringBuilder1 As New StringBuilder
            For num As Integer = position + spansToDelete - 1 To position Step -1
                list1.Add(RemoveSpan(num))
                stringBuilder1.Insert(0, _currentProjectionSnapshot.GetSourceSpan(num).GetText())
            Next

            Dim stringBuilder2 As New StringBuilder
            Dim num2 As Integer = position
            For Each item As ITextSpan In spansToInsert
                AddSpan(Math.Min(Threading.Interlocked.Increment(num2), num2 - 1), item)
                stringBuilder2.Append(item.GetSpan(item.TextBuffer.CurrentSnapshot).GetText())
            Next

            Dim addedBuffers As New List(Of ITextBuffer)
            Dim removedBuffers As New List(Of ITextBuffer)
            _sourceBufferSet.FinishTransaction(addedBuffers, removedBuffers)

            For Each item2 As ITextBuffer In addedBuffers
                AddHandler item2.Changed, AddressOf OnSourceTextChanged
            Next

            For Each item3 As ITextBuffer In removedBuffers
                RemoveHandler item3.Changed, AddressOf OnSourceTextChanged
            Next

            Dim num3 As Integer = 0
            For i As Integer = 0 To position - 1
                num3 += _currentProjectionSnapshot.GetSourceSpan(i).Length
            Next

            Dim j As Integer = 0
            Dim num4 As Integer
            num4 = Math.Min(stringBuilder1.Length, stringBuilder2.Length)
            While j < num4 AndAlso stringBuilder1(j) = stringBuilder2(j)
                j += 1
            End While

            If j > 0 Then
                stringBuilder1.Remove(0, j)
                stringBuilder2.Remove(0, j)
                num3 += j
            End If

            num4 -= j
            j = 0
            Dim length1 As Integer = stringBuilder1.Length
            Dim length2 As Integer
            length2 = stringBuilder2.Length
            While j < num4 AndAlso stringBuilder1(length1 - j - 1) = stringBuilder2(length2 - j - 1)
                j += 1
            End While

            If j > 0 Then
                stringBuilder1.Remove(length1 - j, j)
                stringBuilder2.Remove(length2 - j, j)
            End If

            Dim beforeSnapshot = CType(CurrentProjectionSnapshot, ProjectionSnapshot)
            Dim normalizedTextChanges As NormalizedTextChangeCollection = Nothing

            If stringBuilder1.Length > 0 OrElse stringBuilder2.Length > 0 Then
                Dim list2 As New List(Of TextChange)
                list2.Add(New TextChange(num3, stringBuilder1.ToString(), stringBuilder2.ToString()))
                normalizedTextChanges = New NormalizedTextChangeCollection(list2)
                SetCurrentVersionAndSnapshot(normalizedTextChanges)
            Else
                SetCurrentVersionAndSnapshot(New NormalizedTextChangeCollection(New List(Of TextChange)))
            End If

            RaiseSourceBuffersChangedEvent(addedBuffers, removedBuffers)
            RaiseSourceSpansChangedEvent(spansToInsert, list1)
            If normalizedTextChanges IsNot Nothing Then
                RaiseChangedEvent(beforeSnapshot, CurrentProjectionSnapshot, normalizedTextChanges, Nothing)
            End If
        End Sub

        Private Sub AddSpan(position As Integer, sourceSpan As ITextSpan)
            sourceSpans.Insert(position, sourceSpan)
            _sourceBufferSet.AddSpan(sourceSpan)
        End Sub

        Private Function RemoveSpan(position As Integer) As ITextSpan
            Dim textSpan As ITextSpan = sourceSpans(position)
            sourceSpans.RemoveAt(position)
            _sourceBufferSet.RemoveSpan(textSpan)
            Return textSpan
        End Function

        Private Sub CheckInsertionLegality(spansToInsert As IEnumerable(Of ITextSpan))
            For Each item As ITextSpan In spansToInsert
                If item Is Nothing Then
                    Throw New ArgumentNullException("spansToInsert")
                End If
            Next
        End Sub

        Private Sub OnSourceTextChanged(sender As Object, e As TextChangedEventArgs)
            Dim key As ITextBuffer = CType(sender, ITextBuffer)
            Dim _value As Collections.Generic.List(Of ITextChange) = Nothing
            If Not pendingSourceChanges.TryGetValue(key, _value) Then
                _value = New List(Of ITextChange)
                _value.Add(New TextChange(Integer.MaxValue, "", ""))
                pendingSourceChanges.Add(key, _value)
            End If

            Dim num As Integer = 0
            Dim index As Integer = 0
            For Each change As ITextChange In e.Changes
                While _value(index).Position <= change.Position
                    num += _value(Math.Min(Threading.Interlocked.Increment(index), index - 1)).Delta
                End While
                _value.Insert(index, If((num <> 0), New TextChange(change.Position - num, change.OldText, change.NewText), change))
            Next

            If Not EditTransactionInProgress Then
                Dim normalizedTextChangeCollection1 As NormalizedTextChangeCollection = InterpretSourceChanges()
                If normalizedTextChangeCollection1 IsNot Nothing Then
                    FinishChangeApplication(normalizedTextChangeCollection1, CurrentProjectionSnapshot, e.SourceToken)
                End If
                pendingSourceChanges.Clear()
            End If
        End Sub

        Private Function InterpretSourceChanges() As NormalizedTextChangeCollection
            Dim list1 As New List(Of TextChange)
            For Each pendingSourceChange As KeyValuePair(Of ITextBuffer, List(Of ITextChange)) In pendingSourceChanges
                Dim _value As List(Of ITextChange) = pendingSourceChange.Value
                For i As Integer = 0 To _value.Count - 1 - 1
                    InterpretSourceBufferChange(pendingSourceChange.Key, _value(i), list1)
                Next
            Next

            If list1.Count <= 0 Then
                Return Nothing
            End If
            Return New NormalizedTextChangeCollection(list1)
        End Function

        Private Sub InterpretSourceBufferChange(changedBuffer As ITextBuffer, change As ITextChange, projectedChanges As List(Of TextChange))
            Dim projSnapshot = CType(CurrentProjectionSnapshot, ProjectionSnapshot)
            Dim position As Integer = change.Position
            Dim span As New Span(position, change.OldLength)
            Dim newLength1 As Integer = change.NewLength
            Dim num As Integer = 0
            Dim num2 As Integer = 0

            For Each sourceSpan As ITextSpan In sourceSpans
                Dim trackingMode1 As SpanTrackingMode = sourceSpan.TrackingMode
                Dim span2 As Span = projSnapshot.GetSourceSpan(num2)
                If sourceSpan.TextBuffer Is changedBuffer Then
                    Dim span3 As Span? = span.Overlap(span2)
                    If span3.HasValue AndAlso span3.Value.Length > 0 Then
                        Dim newText1 As String = ""
                        Dim startIndex As Integer = Math.Max(span2.Start - span.Start, 0)
                        Dim oldText1 As String = change.OldText.Substring(startIndex, span3.Value.Length)
                        Dim num3 As Integer = -span3.Value.Length
                        If span3.Value.Start = position AndAlso change.NewLength > 0 AndAlso InsertionLiesInSpan(span2, position, trackingMode1) Then
                            newText1 = change.NewText
                            num3 += newLength1
                        End If
                        projectedChanges.Add(New TextChange(num + span3.Value.Start - span2.Start, oldText1, newText1))

                    ElseIf newLength1 > 0 AndAlso InsertionLiesInSpan(span2, position, trackingMode1) Then
                        projectedChanges.Add(New TextChange(num + position - span2.Start, "", change.NewText))
                    End If
                End If
                num += span2.Length
                num2 += 1
            Next
        End Sub

        Private Shared Function InsertionLiesInSpan(rawSpan As Span, position As Integer, mode As SpanTrackingMode) As Boolean
            Dim flag As Boolean = rawSpan.Contains(position)
            If mode = SpanTrackingMode.EdgeInclusive Then
                If Not flag Then
                    Return position = rawSpan.End
                End If
                Return True
            End If

            If flag Then
                Return position <> rawSpan.Start
            End If

            Return False
        End Function

        Protected Overrides Function ApplyChanges(changes1 As List(Of TextChange), sourceToken1 As Object) As NormalizedTextChangeCollection
            Dim dictionary1 As New Dictionary(Of ITextBuffer, ITextEdit)
            For Each change As TextChange In changes1
                If change.OldLength > 0 Then
                    Dim list1 As IList(Of SnapshotSpan) = CurrentProjectionSnapshot.MapToSourceSnapshots(New Span(change.Position, change.OldLength))
                    For Each item As SnapshotSpan In list1
                        Dim textBuffer1 As ITextBuffer = item.Snapshot.TextBuffer
                        Dim _value As ITextEdit = Nothing
                        If Not dictionary1.TryGetValue(textBuffer1, _value) Then
                            _value = textBuffer1.CreateEdit(sourceToken1)
                            dictionary1.Add(textBuffer1, _value)
                        End If
                        _value.Delete(item)
                    Next
                End If

                If change.NewLength > 0 Then
                    Dim snapshotPoint1 As SnapshotPoint = CurrentProjectionSnapshot.MapToSourceSnapshot(change.Position)
                    Dim textBuffer2 As ITextBuffer = snapshotPoint1.Snapshot.TextBuffer
                    Dim value2 As ITextEdit = Nothing
                    If Not dictionary1.TryGetValue(textBuffer2, value2) Then
                        value2 = textBuffer2.CreateEdit()
                        dictionary1.Add(textBuffer2, value2)
                    End If
                    value2.Insert(snapshotPoint1.Position, change.NewText)
                End If
            Next

            _EditTransactionInProgress = True
            For Each value3 As ITextEdit In dictionary1.Values
                value3.Apply()
                value3.Dispose()
            Next

            Dim result As NormalizedTextChangeCollection = InterpretSourceChanges()
            pendingSourceChanges.Clear()
            _EditTransactionInProgress = False
            Return result
        End Function

        Protected Overrides Function TakeSnapshot() As ITextSnapshot
            Dim list1 As New List(Of SnapshotSpan)
            For Each sourceSpan As ITextSpan In sourceSpans
                list1.Add(sourceSpan.GetSpan(sourceSpan.TextBuffer.CurrentSnapshot))
            Next

            _currentProjectionSnapshot = New ProjectionSnapshot(Me, MyBase.CurrentVersion, list1)
            Return CurrentProjectionSnapshot
        End Function

        Private Sub RaiseSourceBuffersChangedEvent(addedBuffers As IList(Of ITextBuffer), removedBuffers As IList(Of ITextBuffer))
            RaiseEvent SourceBuffersChanged(Me, New BuffersChangedEventArgs(addedBuffers, removedBuffers))
        End Sub

        Private Sub RaiseSourceSpansChangedEvent(spansToInsert As IList(Of ITextSpan), deletedSpans As List(Of ITextSpan))
            RaiseEvent SourceSpansChanged(Me, New SpansChangedEventArgs(spansToInsert, deletedSpans))
        End Sub
    End Class
End Namespace
