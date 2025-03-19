Imports System.Collections.Generic

Namespace Microsoft.Nautilus.Text.Projection.Implementation
    Friend Class BufferGraph
        Implements IBufferGraph

        Private targetBufferMap As New Dictionary(Of ITextBuffer, List(Of IProjectionBuffer))

        Public ReadOnly Property TopBuffer As ITextBuffer Implements IBufferGraph.TopBuffer


        Public Sub New(topBuffer As ITextBuffer)
            If topBuffer Is Nothing Then
                Throw New ArgumentNullException("_topBuffer")
            End If

            _TopBuffer = topBuffer
            targetBufferMap.Add(topBuffer, Nothing)
            If TypeOf topBuffer IsNot IProjectionBuffer Then Return

            Dim projectionBuffer = CType(topBuffer, IProjectionBuffer)

            AddHandler projectionBuffer.SourceBuffersChanged, AddressOf SourceBuffersChanged
            Dim sourceBuffers1 As IList(Of ITextBuffer) = projectionBuffer.SourceBuffers
            For Each item As ITextBuffer In sourceBuffers1
                AddSourceBuffer(projectionBuffer, item)
            Next
        End Sub

        Public Function MapDownToBuffer(topPosition As Integer, targetBuffer As ITextBuffer) As SnapshotPoint Implements IBufferGraph.MapDownToBuffer
            If targetBuffer Is Nothing Then
                Throw New ArgumentNullException("targetBuffer")
            End If

            If topPosition < 0 OrElse topPosition > _TopBuffer.CurrentSnapshot.Length Then
                Throw New ArgumentOutOfRangeException("topPosition")
            End If

            If Not targetBufferMap.ContainsKey(targetBuffer) Then
                Throw New ArgumentException("specified targetBuffer is not a member of buffer graph")
            End If

            Dim position As Integer = topPosition
            Dim textBuffer As ITextBuffer = _TopBuffer
            Dim textSnapshot As ITextSnapshot = _TopBuffer.CurrentSnapshot

            While textBuffer IsNot targetBuffer
                If TypeOf textBuffer IsNot IProjectionBuffer Then
                    Return Nothing
                End If

                Dim projectionBuffer = CType(textBuffer, IProjectionBuffer)
                Dim snapshotPoint = projectionBuffer.CurrentProjectionSnapshot.MapToSourceSnapshot(position)
                position = snapshotPoint.Position
                textSnapshot = snapshotPoint.Snapshot
                textBuffer = textSnapshot.TextBuffer
            End While

            Return New SnapshotPoint(textSnapshot, position)
        End Function

        Public Function MapUpToTop(point As SnapshotPoint) As SnapshotPoint? Implements IBufferGraph.MapUpToTop
            If Not targetBufferMap.ContainsKey(point.Snapshot.TextBuffer) Then
                Throw New ArgumentException("specified SnapshotPoint does not belong to a TextBuffer that is a member of buffer graph")
            End If

            Return MapUpToTopGuts(point)
        End Function

        Private Function MapUpToTopGuts(point As SnapshotPoint) As SnapshotPoint?
            If point.Snapshot.TextBuffer Is _TopBuffer Then
                Return point
            End If

            Dim projectionBuffers = targetBufferMap(point.Snapshot.TextBuffer)
            For Each projBuffer In projectionBuffers
                Dim snapshotPoint = projBuffer.CurrentProjectionSnapshot.MapFromSourceSnapshot(point)
                If snapshotPoint.HasValue Then
                    snapshotPoint = MapUpToTopGuts(snapshotPoint.Value)
                    If snapshotPoint.HasValue Then Return snapshotPoint
                End If
            Next

            Return Nothing
        End Function

        Public Function MapDownToBuffer(topSpan As Span, targetBuffer As ITextBuffer) As NormalizedSpanCollection Implements IBufferGraph.MapDownToBuffer
            If targetBuffer Is Nothing Then
                Throw New ArgumentNullException("targetBuffer")
            End If

            If Not targetBufferMap.ContainsKey(targetBuffer) Then
                Throw New ArgumentException("specified targetBuffer is not a member of buffer graph")
            End If

            If topSpan.End > _TopBuffer.CurrentSnapshot.Length Then
                Throw New ArgumentOutOfRangeException("topSpan")
            End If

            Dim targetSpans As New List(Of Span)
            If targetBuffer Is _TopBuffer Then
                targetSpans.Add(topSpan)
            Else
                Dim list1 As New List(Of SnapshotSpan)
                list1.Add(New SnapshotSpan(_TopBuffer.CurrentSnapshot, topSpan))

                Do
                    list1 = MapDownOneLevel(list1, targetBuffer, targetSpans)
                Loop While list1.Count > 0
            End If
            Return New NormalizedSpanCollection(targetSpans)
        End Function

        Private Function MapDownOneLevel(
                              inputSpans As List(Of SnapshotSpan),
                              targetBuffer As ITextBuffer,
                              ByRef targetSpans As List(Of Span)
                     ) As List(Of SnapshotSpan)

            Dim results As New List(Of SnapshotSpan)
            For Each inputSpan As SnapshotSpan In inputSpans
                Dim projectionBuffer = CType(inputSpan.Snapshot.TextBuffer, IProjectionBuffer)
                Dim snapshotSpans = projectionBuffer.CurrentProjectionSnapshot.MapToSourceSnapshots(inputSpan)
                For Each snapshotSpan In snapshotSpans
                    Dim textBuffer = snapshotSpan.Snapshot.TextBuffer
                    If textBuffer Is targetBuffer Then
                        targetSpans.Add(snapshotSpan.Span)
                    ElseIf TypeOf textBuffer Is IProjectionBuffer Then
                        results.Add(snapshotSpan)
                    End If
                Next
            Next
            Return results
        End Function

        Public Function MapUpToTop(span As SnapshotSpan) As NormalizedSpanCollection Implements IBufferGraph.MapUpToTop
            If Not targetBufferMap.ContainsKey(span.Snapshot.TextBuffer) Then
                Throw New ArgumentException("specified SnapshotSpan does not belong to a TextBuffer that is a member of buffer graph")
            End If

            If span.Snapshot.TextBuffer Is _TopBuffer Then
                Return New NormalizedSpanCollection(span)
            End If

            Dim topSpans As New List(Of Span)
            Dim snapshotSpans As New List(Of SnapshotSpan)(1)
            snapshotSpans.Add(span)

            Do
                snapshotSpans = MapUpOneLevel(snapshotSpans, topSpans)
            Loop While snapshotSpans.Count > 0

            Return New NormalizedSpanCollection(topSpans)
        End Function

        Private Function MapUpOneLevel(spans As List(Of SnapshotSpan), ByRef topSpans As List(Of Span)) As List(Of SnapshotSpan)
            Dim topSnapshotSpans As New List(Of SnapshotSpan)
            For Each span In spans
                Dim projectionBuffers As List(Of IProjectionBuffer) = Nothing
                If Not targetBufferMap.TryGetValue(span.Snapshot.TextBuffer, projectionBuffers) Then
                    Continue For
                End If

                For Each projectionBuffer In projectionBuffers
                    Dim topSpans2 = projectionBuffer.CurrentProjectionSnapshot.MapFromSourceSnapshot(span)
                    If projectionBuffer Is _TopBuffer Then
                        topSpans.AddRange(topSpans2)
                        Continue For
                    End If

                    For Each topSpan As Span In topSpans2
                        topSnapshotSpans.Add(New SnapshotSpan(projectionBuffer.CurrentSnapshot, topSpan))
                    Next
                Next
            Next

            Return topSnapshotSpans
        End Function

        Private Sub SourceBuffersChanged(sender As Object, e As BuffersChangedEventArgs)
            Dim projBuffer = CType(sender, IProjectionBuffer)
            For Each addedBuffer In e.AddedBuffers
                AddSourceBuffer(projBuffer, addedBuffer)
            Next

            For Each removedBuffer In e.RemovedBuffers
                RemoveSourceBuffer(projBuffer, removedBuffer)
            Next
        End Sub

        Private Sub AddSourceBuffer(projBuffer As IProjectionBuffer, addedBuffer As ITextBuffer)
            Dim flag As Boolean = False
            Dim projectionBuffers As Collections.Generic.List(Of IProjectionBuffer) = Nothing
            If Not targetBufferMap.TryGetValue(addedBuffer, projectionBuffers) Then
                projectionBuffers = New List(Of IProjectionBuffer)
                targetBufferMap.Add(addedBuffer, projectionBuffers)
                flag = True
            End If

            projectionBuffers.Add(projBuffer)
            If Not TypeOf addedBuffer Is IProjectionBuffer Then Return

            Dim projectionBuffer = CType(addedBuffer, IProjectionBuffer)
            If flag Then
                AddHandler projectionBuffer.SourceBuffersChanged, AddressOf SourceBuffersChanged
            End If

            For Each sourceBuffer In projectionBuffer.SourceBuffers
                AddSourceBuffer(projectionBuffer, sourceBuffer)
            Next
        End Sub

        Private Sub RemoveSourceBuffer(projBuffer As IProjectionBuffer, removedBuffer As ITextBuffer)
            Dim projectionBuffers = targetBufferMap(removedBuffer)
            projectionBuffers.Remove(projBuffer)
            If projectionBuffers.Count <> 0 Then Return

            targetBufferMap.Remove(removedBuffer)
            If Not TypeOf removedBuffer Is IProjectionBuffer Then Return

            Dim projectionBuffer = CType(removedBuffer, IProjectionBuffer)
            RemoveHandler projectionBuffer.SourceBuffersChanged, AddressOf SourceBuffersChanged

            For Each sourceBuffer In projectionBuffer.SourceBuffers
                RemoveSourceBuffer(projectionBuffer, sourceBuffer)
            Next
        End Sub
    End Class
End Namespace
