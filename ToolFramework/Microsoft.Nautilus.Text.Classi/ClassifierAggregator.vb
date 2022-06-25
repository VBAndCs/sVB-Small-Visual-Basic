Imports System.Collections.Generic
Imports System.ComponentModel.Composition
Imports Microsoft.Nautilus.Core

Namespace Microsoft.Nautilus.Text.Classification
    Friend Class ClassifierAggregator
        Implements IClassifier

        Private Class Point
            Private _owner As ClassificationSpan
            Private _point As Integer

            Public ReadOnly Property ClassificationSpan As ClassificationSpan
                Get
                    Return _owner
                End Get
            End Property

            Public ReadOnly Property Value As Integer
                Get
                    Return _point
                End Get
            End Property

            Public Sub New(owner As ClassificationSpan, point1 As Integer)
                _owner = owner
                _point = point1
            End Sub
        End Class

        Private _textBuffer As ITextBuffer
        Private _classifiers As List(Of IClassifier)
        Private _classificationTypeRegistry As IClassificationTypeRegistry
        Private _classifierProviders As IEnumerable(Of ImportInfo(Of Func(Of ITextBuffer, IClassifier), IContentTypeMetadata))

        Public Event ClassificationChanged As EventHandler(Of ClassificationChangedEventArgs) Implements IClassifier.ClassificationChanged

        Friend Sub New(textBuffer As ITextBuffer, providers As IEnumerable(Of ImportInfo(Of Func(Of ITextBuffer, IClassifier), IContentTypeMetadata)), classificationTypeRegistry As IClassificationTypeRegistry)
            If textBuffer Is Nothing Then
                Throw New ArgumentNullException("textBuffer")
            End If

            If providers Is Nothing Then
                Throw New ArgumentNullException("providers")
            End If

            If classificationTypeRegistry Is Nothing Then
                Throw New ArgumentNullException("classificationTypeRegistry")
            End If

            _textBuffer = textBuffer
            _classificationTypeRegistry = classificationTypeRegistry
            _classifierProviders = providers
            _classifiers = New List(Of IClassifier)
            BuildClassifierList()
            AddHandler _textBuffer.ContentTypeChanged, Sub() BuildClassifierList()
        End Sub

        Public Function GetClassificationSpans(span As SnapshotSpan) As IList(Of ClassificationSpan) Implements IClassifier.GetClassificationSpans
            If _classifiers.Count = 1 Then
                Return _classifiers(0).GetClassificationSpans(span)
            End If

            Dim list1 As New List(Of ClassificationSpan)
            Dim num As Integer = 0
            For Each classifier As IClassifier In _classifiers
                Dim classificationSpans As IList(Of ClassificationSpan) = classifier.GetClassificationSpans(span)
                If classificationSpans.Count > 0 Then
                    num += 1
                    list1.AddRange(classificationSpans)
                End If
            Next

            If num = 1 Then
                Return list1
            End If
            Return AggregateClassificationSpans(span, list1)
        End Function

        Private Sub OnClassificationChanged(sender As Object, e As ClassificationChangedEventArgs)
            RaiseEvent ClassificationChanged(sender, e)
        End Sub

        Private Sub BuildClassifierList()
            For Each classifier3 As IClassifier In _classifiers
                RemoveHandler classifier3.ClassificationChanged, AddressOf OnClassificationChanged
            Next

            _classifiers.Clear()
            For Each classifierProvider As ImportInfo(Of Func(Of ITextBuffer, IClassifier), IContentTypeMetadata) In _classifierProviders
                For Each contentType As String In classifierProvider.Metadata.ContentTypes
                    If ContentTypeHelper.IsSame(_textBuffer.ContentType, contentType) Then
                        Dim classifier As IClassifier = classifierProvider.GetBoundValue()(_textBuffer)
                        If classifier IsNot Nothing Then
                            _classifiers.Add(classifier)
                            AddHandler classifier.ClassificationChanged, AddressOf OnClassificationChanged
                        End If
                        Return
                    End If
                Next
            Next

            For Each classifierProvider2 As ImportInfo(Of Func(Of ITextBuffer, IClassifier), IContentTypeMetadata) In _classifierProviders
                For Each contentType2 As String In classifierProvider2.Metadata.ContentTypes
                    If ContentTypeHelper.IsOfType(_textBuffer.ContentType, contentType2) Then
                        Dim classifier2 As IClassifier = classifierProvider2.GetBoundValue()(_textBuffer)
                        If classifier2 IsNot Nothing Then
                            _classifiers.Add(classifier2)
                            AddHandler classifier2.ClassificationChanged, AddressOf OnClassificationChanged
                        End If
                        Return
                    End If
                Next
            Next
        End Sub

        Private Shared Sub AddPoint(sortedPoints As SortedDictionary(Of Integer, IList(Of Point)), classification As ClassificationSpan, point1 As Integer)
            If sortedPoints.ContainsKey(point1) Then
                sortedPoints(point1).Add(New Point(classification, point1))
                Return
            End If

            Dim list1 As New List(Of Point)
            list1.Add(New Point(classification, point1))
            sortedPoints.Add(point1, list1)
        End Sub

        Private Function AggregateClassificationSpans(requestedRange As SnapshotSpan, spans As IList(Of ClassificationSpan)) As IList(Of ClassificationSpan)
            Dim sortedDictionary1 As New SortedDictionary(Of Integer, IList(Of Point))
            For Each span2 As ClassificationSpan In spans
                Dim span As Span = span2.GetSpan(requestedRange.Snapshot)
                If requestedRange.Span.OverlapsWith(span) Then
                    AddPoint(sortedDictionary1, span2, span.Start)
                    AddPoint(sortedDictionary1, span2, span.End)
                End If
            Next

            Dim list1 As New List(Of ClassificationSpan)
            Dim dictionary1 As New Dictionary(Of ClassificationSpan, IClassificationType)
            Dim enumerator2 As IEnumerator(Of IList(Of Point)) = sortedDictionary1.Values.GetEnumerator()
            Dim val As Integer = 0
            While enumerator2.MoveNext()
                Dim num As Integer = enumerator2.Current(0).Value
                If dictionary1.Count > 0 Then
                    Dim list2 As New List(Of IClassificationType)
                    For Each value1 As IClassificationType In dictionary1.Values
                        If Not list2.Contains(value1) Then
                            list2.Add(value1)
                        End If
                    Next

                    Dim classificationType As IClassificationType = Nothing
                    classificationType = (If((list2.Count <> 1), _classificationTypeRegistry.CreateTransientClassificationType(list2), list2(0)))
                    val = Math.Max(val, requestedRange.Span.Start)
                    num = Math.Min(num, requestedRange.Span.End)
                    list1.Add(New ClassificationSpan(requestedRange.Snapshot.CreateTextSpan(val, num - val, SpanTrackingMode.EdgeExclusive), classificationType))
                End If

                For Each item As Point In enumerator2.Current
                    If dictionary1.ContainsKey(item.ClassificationSpan) Then
                        dictionary1.Remove(item.ClassificationSpan)
                    Else
                        dictionary1.Add(item.ClassificationSpan, item.ClassificationSpan.ClassificationType)
                    End If
                Next

                val = num
            End While

            Return list1
        End Function
    End Class
End Namespace
