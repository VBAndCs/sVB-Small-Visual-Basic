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
                Dim snapshotPoint1 As SnapshotPoint = projectionBuffer.CurrentProjectionSnapshot.MapToSourceSnapshot(position)
                position = snapshotPoint1.Position
                textSnapshot = snapshotPoint1.Snapshot
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

            Dim list1 As List(Of IProjectionBuffer) = targetBufferMap(point.Snapshot.TextBuffer)
            For Each item As IProjectionBuffer In list1
                Dim snapshotPoint1 = item.CurrentProjectionSnapshot.MapFromSourceSnapshot(point)
                If snapshotPoint1.HasValue Then
                    snapshotPoint1 = MapUpToTopGuts(snapshotPoint1.Value)
                    If snapshotPoint1.HasValue Then
                        Return snapshotPoint1
                    End If
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

        Private Function MapDownOneLevel(inputSpans As List(Of SnapshotSpan), targetBuffer As ITextBuffer, ByRef targetSpans As List(Of Span)) As List(Of SnapshotSpan)
            Dim list1 As New List(Of SnapshotSpan)
            For Each inputSpan As SnapshotSpan In inputSpans
                Dim projectionBuffer As IProjectionBuffer = CType(inputSpan.Snapshot.TextBuffer, IProjectionBuffer)
                Dim list2 As IList(Of SnapshotSpan) = projectionBuffer.CurrentProjectionSnapshot.MapToSourceSnapshots(inputSpan)
                For Each item As SnapshotSpan In list2
                    Dim textBuffer As ITextBuffer = item.Snapshot.TextBuffer
                    If textBuffer Is targetBuffer Then
                        targetSpans.Add(item.Span)
                    ElseIf TypeOf textBuffer Is IProjectionBuffer Then
                        list1.Add(item)
                    End If
                Next
            Next
            Return list1
        End Function

        Public Function MapUpToTop(span1 As SnapshotSpan) As NormalizedSpanCollection Implements IBufferGraph.MapUpToTop
            If Not targetBufferMap.ContainsKey(span1.Snapshot.TextBuffer) Then
                Throw New ArgumentException("specified SnapshotSpan does not belong to a TextBuffer that is a member of buffer graph")
            End If

            If span1.Snapshot.TextBuffer Is _TopBuffer Then
                Return New NormalizedSpanCollection(span1)
            End If

            Dim topSpans As New List(Of Span)
            Dim list1 As New List(Of SnapshotSpan)(1)
            list1.Add(span1)

            Do
                list1 = MapUpOneLevel(list1, topSpans)
            Loop While list1.Count > 0

            Return New NormalizedSpanCollection(topSpans)
        End Function

        Private Function MapUpOneLevel(spans As List(Of SnapshotSpan), ByRef topSpans As List(Of Span)) As List(Of SnapshotSpan)
            Dim list1 As New List(Of SnapshotSpan)
            For Each span1 As SnapshotSpan In spans
                Dim value1 As Collections.Generic.List(Of IProjectionBuffer) = Nothing
                If Not targetBufferMap.TryGetValue(span1.Snapshot.TextBuffer, value1) Then
                    Continue For
                End If
                For Each item As IProjectionBuffer In value1
                    Dim list2 As IList(Of Span) = item.CurrentProjectionSnapshot.MapFromSourceSnapshot(span1)
                    If item Is _TopBuffer Then
                        topSpans.AddRange(list2)
                        Continue For
                    End If

                    For Each item2 As Span In list2
                        list1.Add(New SnapshotSpan(item.CurrentSnapshot, item2))
                    Next
                Next
            Next
            Return list1
        End Function

        Private Sub SourceBuffersChanged(sender As Object, e As BuffersChangedEventArgs)
            Dim projBuffer As IProjectionBuffer = CType(sender, IProjectionBuffer)
            For Each addedBuffer As ITextBuffer In e.AddedBuffers
                AddSourceBuffer(projBuffer, addedBuffer)
            Next

            For Each removedBuffer As ITextBuffer In e.RemovedBuffers
                RemoveSourceBuffer(projBuffer, removedBuffer)
            Next
        End Sub

        Private Sub AddSourceBuffer(projBuffer As IProjectionBuffer, addedBuffer As ITextBuffer)
            Dim flag As Boolean = False
            Dim value1 As Collections.Generic.List(Of IProjectionBuffer) = Nothing
            If Not targetBufferMap.TryGetValue(addedBuffer, value1) Then
                value1 = New List(Of IProjectionBuffer)
                targetBufferMap.Add(addedBuffer, value1)
                flag = True
            End If

            value1.Add(projBuffer)

            If Not TypeOf addedBuffer Is IProjectionBuffer Then Return
            Dim projectionBuffer = CType(addedBuffer, IProjectionBuffer)

            If flag Then
                AddHandler projectionBuffer.SourceBuffersChanged, AddressOf SourceBuffersChanged
            End If

            For Each sourceBuffer As ITextBuffer In projectionBuffer.SourceBuffers
                AddSourceBuffer(projectionBuffer, sourceBuffer)
            Next
        End Sub

        Private Sub RemoveSourceBuffer(projBuffer As IProjectionBuffer, removedBuffer As ITextBuffer)
            Dim list1 As IList(Of IProjectionBuffer) = targetBufferMap(removedBuffer)
            list1.Remove(projBuffer)
            If list1.Count <> 0 Then
                Return
            End If

            targetBufferMap.Remove(removedBuffer)

            If Not TypeOf removedBuffer Is IProjectionBuffer Then Return
            Dim projectionBuffer = CType(removedBuffer, IProjectionBuffer)

            RemoveHandler projectionBuffer.SourceBuffersChanged, AddressOf SourceBuffersChanged
            For Each sourceBuffer As ITextBuffer In projectionBuffer.SourceBuffers
                RemoveSourceBuffer(projectionBuffer, sourceBuffer)
            Next
        End Sub
    End Class
End Namespace
