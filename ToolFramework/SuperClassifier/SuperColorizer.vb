Imports System.Collections.Generic
Imports Microsoft.Nautilus.Text
Imports Microsoft.Nautilus.Text.Classification

Namespace SuperClassifier
    Friend Class SuperColorizer
        Implements IClassifier

        Private classificationTypeRegistry As IClassificationTypeRegistry
        Private mruCache As New List(Of TokenCache)
        Private tokenizer As Tokenizer

        Public Property Buffer As ITextBuffer

        Public Event ClassificationChanged As EventHandler(Of ClassificationChangedEventArgs) Implements IClassifier.ClassificationChanged

        Public Sub New(textBuffer As ITextBuffer, classificationTypeRegistry As IClassificationTypeRegistry, tokenizer As Tokenizer)
            Me.classificationTypeRegistry = classificationTypeRegistry
            _Buffer = textBuffer
            Me.tokenizer = tokenizer

            AddHandler _Buffer.Changed, AddressOf OnBufferChanged
            AddHandler _Buffer.ContentTypeChanged, AddressOf OnBufferContentTypeChanged
        End Sub

        Private Sub OnBufferContentTypeChanged(sender As Object, e As EventArgs)
            mruCache.Clear()
            RemoveHandler _Buffer.Changed, AddressOf OnBufferChanged
            RemoveHandler _Buffer.ContentTypeChanged, AddressOf OnBufferContentTypeChanged
        End Sub

        Private Sub OnBufferChanged(sender As Object, e As TextChangedEventArgs)
            Dim invalidatedSpan As ITextSpan = tokenizer.GetInvalidatedSpan(e.Changes)
            If Not invalidatedSpan.GetSpan(e.After).IsEmpty Then
                RaiseEvent ClassificationChanged(Me, New ClassificationChangedEventArgs(invalidatedSpan))
            End If
            mruCache.Clear()
        End Sub

        Private Function GetTokensCovering(start As Integer, length As Integer) As List(Of Token)
            For i As Integer = 0 To mruCache.Count - 1
                Dim tokenCache1 As TokenCache = mruCache(i)
                If tokenCache1.startIndex = start AndAlso tokenCache1.length = length Then
                    mruCache.RemoveAt(i)
                    mruCache.Insert(0, tokenCache1)
                    Return tokenCache1.tokensCoveringCache
                End If
            Next

            Dim tokensCovering As List(Of Token) = tokenizer.GetTokensCovering(start, length)
            If mruCache.Count > 100 Then
                mruCache(mruCache.Count - 1) = New TokenCache(start, length, tokensCovering)
            Else
                mruCache.Add(New TokenCache(start, length, tokensCovering))
            End If

            Return tokensCovering
        End Function

        Public Function GetClassificationSpans(span As SnapshotSpan) As IList(Of ClassificationSpan) Implements IClassifier.GetClassificationSpans
            Dim span2 = span.Span
            Dim classificationSpans As New List(Of ClassificationSpan)
            Dim tokensCovering As List(Of Token) = GetTokensCovering(span2.Start, span2.Length)

            For Each token In tokensCovering
                If token.TokenEnd > span2.Start Then
                    If token.TokenStart >= span2.End Then Return classificationSpans

                    Dim textSpan = New TextSpan(_Buffer.CurrentSnapshot, token.TokenStart, token.TokenLength, SpanTrackingMode.EdgeExclusive)
                    Dim classificationType = classificationTypeRegistry.GetClassificationType(token.Classification)
                    classificationSpans.Add(New ClassificationSpan(textSpan, classificationType))
                End If
            Next

            Return classificationSpans
        End Function
    End Class
End Namespace
