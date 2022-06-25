Imports System.Collections.Generic
Imports Microsoft.Nautilus.Text.Projection

Namespace Microsoft.Nautilus.Text.Classification
    Friend Class ProjectionClassifier
        Implements IClassifier

        Public Delegate Function AggregatorProvider(projBuffer As ITextBuffer) As IClassifier
        Private projBuffer As IProjectionBuffer
        Private getAggregator As AggregatorProvider
        Private queuedChanges As New List(Of ClassificationChangedEventArgs)

        Public Event ClassificationChanged As EventHandler(Of ClassificationChangedEventArgs) Implements IClassifier.ClassificationChanged

        Friend Sub New(projBuffer As IProjectionBuffer, getAggregator As AggregatorProvider)
            If projBuffer Is Nothing Then
                Throw New ArgumentNullException("projBuffer")
            End If

            If getAggregator Is Nothing Then
                Throw New ArgumentNullException("getAggregator")
            End If

            Me.projBuffer = projBuffer
            Me.getAggregator = getAggregator
            AddHandler projBuffer.Changed, AddressOf ProjectionContentChanged
            AddHandler projBuffer.SourceBuffersChanged, AddressOf SourceBuffersChanged
            For Each sourceSnapshot As ITextSnapshot In projBuffer.CurrentProjectionSnapshot.SourceSnapshots
                Dim classifier As IClassifier = Me.getAggregator(sourceSnapshot.TextBuffer)
                AddHandler classifier.ClassificationChanged, AddressOf SourceClassificationChanged
            Next
        End Sub

        Public Function GetClassificationSpans(requestSpan As SnapshotSpan) As IList(Of ClassificationSpan) Implements IClassifier.GetClassificationSpans
            If requestSpan.Snapshot IsNot projBuffer.CurrentSnapshot Then
                Throw New ArgumentException("span applies to wrong or out-of-sync text buffer")
            End If

            Dim projectionSnapshot As IProjectionSnapshot = CType(requestSpan.Snapshot, IProjectionSnapshot)
            Dim list1 As IList(Of SnapshotSpan) = projectionSnapshot.MapToSourceSnapshots(requestSpan)
            Dim list2 As New List(Of ClassificationSpan)

            For Each item As SnapshotSpan In list1
                Dim classifier As IClassifier = getAggregator(item.Snapshot.TextBuffer)
                Dim classificationSpans As IList(Of ClassificationSpan) = classifier.GetClassificationSpans(item)
                For Each item2 As ClassificationSpan In classificationSpans
                    Dim span As SnapshotSpan = item2.GetSpan(item.Snapshot)
                    Dim span2 = span.Overlap(item)
                    If Not span2.HasValue Then
                        Continue For
                    End If

                    Dim span3 As New SnapshotSpan(span.Snapshot, span2.Value)
                    Dim list3 As IList(Of Span) = projectionSnapshot.MapFromSourceSnapshot(span3)
                    For Each item3 As Span In list3
                        list2.Add(New ClassificationSpan(projectionSnapshot.CreateTextSpan(item3, SpanTrackingMode.EdgeExclusive), item2.ClassificationType))
                    Next
                Next
            Next

            Return list2
        End Function

        Private Sub ProjectionContentChanged(sender As Object, e As TextChangedEventArgs)
            If queuedChanges.Count > 0 Then
                If ClassificationChangedEvent IsNot Nothing Then
                    For Each queuedChange As ClassificationChangedEventArgs In queuedChanges
                        PropagateEvents(ClassificationChangedEvent, queuedChange.ChangeSpan)
                    Next
                End If
            End If
            queuedChanges.Clear()
        End Sub

        Private Sub SourceClassificationChanged(sender As Object, e As ClassificationChangedEventArgs)
            If ClassificationChangedEvent IsNot Nothing Then
                If projBuffer.EditTransactionInProgress Then
                    queuedChanges.Add(e)
                Else
                    PropagateEvents(ClassificationChangedEvent, e.ChangeSpan)
                End If
            End If
        End Sub

        Private Sub PropagateEvents(handler As EventHandler(Of ClassificationChangedEventArgs), changeSpan1 As ITextSpan)
            Dim currentProjectionSnapshot1 As IProjectionSnapshot = projBuffer.CurrentProjectionSnapshot
            Dim matchingSnapshot As ITextSnapshot = currentProjectionSnapshot1.GetMatchingSnapshot(changeSpan1.TextBuffer)
            Dim collection As ICollection(Of Span) = currentProjectionSnapshot1.MapFromSourceSnapshot(New SnapshotSpan(matchingSnapshot, changeSpan1.GetSpan(matchingSnapshot)))
            For Each item As Span In collection
                handler(Me, New ClassificationChangedEventArgs(currentProjectionSnapshot1.CreateTextSpan(item, SpanTrackingMode.EdgeExclusive)))
            Next
        End Sub

        Private Sub SourceBuffersChanged(sender As Object, e As BuffersChangedEventArgs)
            For Each removedBuffer As ITextBuffer In e.RemovedBuffers
                Dim classifier As IClassifier = getAggregator(removedBuffer)
                RemoveHandler classifier.ClassificationChanged, AddressOf SourceClassificationChanged
            Next

            For Each addedBuffer As ITextBuffer In e.AddedBuffers
                Dim classifier2 As IClassifier = getAggregator(addedBuffer)
                AddHandler classifier2.ClassificationChanged, AddressOf SourceClassificationChanged
            Next
        End Sub
    End Class
End Namespace
