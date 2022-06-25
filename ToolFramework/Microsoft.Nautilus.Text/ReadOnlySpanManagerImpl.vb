Imports System.Collections.Generic

Namespace Microsoft.Nautilus.Text
    Friend Class ReadOnlySpanManagerImpl
        Inherits ReadOnlyRegionManager

        Private _buffer As ITextBuffer
        Private _readOnlySpanTokens As New List(Of IReadOnlyRegionHandle)
        Private _readOnlySpans As New List(Of ReadOnlySpan)

        Friend Sub New(buffer As ITextBuffer)
            _buffer = buffer
        End Sub

        Public Overrides Function CreateReadOnlyRegion(span1 As Span, trackingMode As SpanTrackingMode, edgeInsertionMode1 As EdgeInsertionMode) As IReadOnlyRegionHandle
            Dim currentSnapshot1 = _buffer.CurrentSnapshot

            If span1.End > currentSnapshot1.Length Then
                Throw New ArgumentOutOfRangeException("span")
            End If

            Dim readOnlyRegionHandle As IReadOnlyRegionHandle = CreateReadOnlyRegion(currentSnapshot1, span1, trackingMode, edgeInsertionMode1)
            Dim newReadOnlySpan As New ReadOnlySpan(currentSnapshot1, span1, trackingMode, edgeInsertionMode1, edgeInsertionMode1)
            Dim list1 As List(Of ReadOnlySpan) = _readOnlySpans.FindAll(Function(s As ReadOnlySpan) s.Intersects(newReadOnlySpan))

            If list1.Count = 0 Then
                _readOnlySpans.Add(newReadOnlySpan)
            Else
                For Each item As ReadOnlySpan In list1
                    If item.TrackingMode = trackingMode Then
                        newReadOnlySpan = MergeReadOnlySpans(item, newReadOnlySpan)
                        _readOnlySpans.Remove(item)
                    End If
                Next

                _readOnlySpans.Add(newReadOnlySpan)
            End If

            _readOnlySpanTokens.Add(readOnlyRegionHandle)
            Return readOnlyRegionHandle
        End Function

        Public Overrides Function GetReadOnlyExtents(span1 As Span) As IList(Of Span)
            Dim currentSnapshot1 As ITextSnapshot = _buffer.CurrentSnapshot
            If span1.End > currentSnapshot1.Length Then
                Throw New ArgumentOutOfRangeException("span")
            End If

            Dim list1 As List(Of ReadOnlySpan) = _readOnlySpans.FindAll(Function(s As ReadOnlySpan) span1.OverlapsWith(s.GetSpan(currentSnapshot1)))
            list1.Sort(Function(first As ReadOnlySpan, second As ReadOnlySpan) As Integer
                           Dim span3 As Span = first.GetSpan(currentSnapshot1)
                           Dim span4 As Span = second.GetSpan(currentSnapshot1)
                           Return span3.Start.CompareTo(span4.Start)
                       End Function)

            Dim list2 As New List(Of Span)
            While list1.Count > 0
                Dim item As Span = list1(0).GetSpan(currentSnapshot1)
                If list1.Count = 1 Then
                    list2.Add(item)
                    list1.Remove(list1(0))
                    Continue While
                End If

                Dim span2 As Span = list1(1).GetSpan(currentSnapshot1)
                If item.End <= span2.Start Then
                    list2.Add(item)
                    list1.Remove(list1(0))
                    Continue While
                End If

                Dim [end] As Integer = item.End
                Dim num As Integer = 1
                While [end] > span2.Start
                    [end] = span2.End
                    num += 1
                    If num >= list1.Count Then
                        Exit While
                    End If
                    span2 = list1(num).GetSpan(currentSnapshot1)
                End While

                list2.Add(New Span(item.Start, span2.End - item.Start))
                list1.RemoveRange(0, num)
            End While

            Return list2
        End Function

        Protected Overrides Function CreateReadOnlyRegion(snapshot As ITextSnapshot, span1 As Span, trackingMode As SpanTrackingMode, edgeInsertionMode1 As EdgeInsertionMode) As IReadOnlyRegionHandle
            If snapshot Is Nothing Then
                Throw New ArgumentNullException("snapshot")
            End If

            Dim readOnlyRegionHandle As IReadOnlyRegionHandle = New ReadOnlyRegionHandle(snapshot, span1, trackingMode, edgeInsertionMode1, Me)
            OnReadOnlyRegionCreated(readOnlyRegionHandle)
            Return readOnlyRegionHandle
        End Function

        Protected Overrides Sub RemoveReadOnlyRegion(readOnlySpanToken As IReadOnlyRegionHandle)
            If readOnlySpanToken Is Nothing Then
                Throw New ArgumentNullException("readOnlySpanToken")
            End If

            _readOnlySpanTokens.Remove(readOnlySpanToken)
            Dim currentSnapshot1 As ITextSnapshot = _buffer.CurrentSnapshot
            Dim tokenSpan As SnapshotSpan = readOnlySpanToken.Span.GetSpan(currentSnapshot1)
            Dim list1 As List(Of IReadOnlyRegionHandle) = _readOnlySpanTokens.FindAll(Function(tok As IReadOnlyRegionHandle) tokenSpan.OverlapsWith(tok.Span.GetSpan(currentSnapshot1)))
            Dim tokenReadOnlySpan As New ReadOnlySpan(currentSnapshot1, tokenSpan, readOnlySpanToken.Span.TrackingMode, readOnlySpanToken.EdgeInsertionMode, readOnlySpanToken.EdgeInsertionMode)

            If list1.Count = 0 Then
                _readOnlySpans.RemoveAll(Function(s As ReadOnlySpan) s.Intersects(tokenReadOnlySpan))
            Else
                Dim list2 As New List(Of ReadOnlySpan)
                For Each item As IReadOnlyRegionHandle In list1
                    Dim extent As ReadOnlySpan = GetExtentOfReadOnlySpan(item.Span.GetSpan(currentSnapshot1))
                    Dim list3 As List(Of ReadOnlySpan) = list2.FindAll(Function(s As ReadOnlySpan) s.Intersects(extent))
                    If list3.Count > 0 Then
                        Dim readOnlySpan1 As ReadOnlySpan = list3(0)
                        list2.Remove(readOnlySpan1)
                        list2.Add(MergeReadOnlySpans(readOnlySpan1, extent))
                    Else
                        list2.Add(extent)
                    End If
                Next

                _readOnlySpans.RemoveAll(Function(s As ReadOnlySpan) tokenSpan.OverlapsWith(s.GetSpan(currentSnapshot1)))
                _readOnlySpans.AddRange(list2)
            End If

            OnReadOnlyRegionRemoved(readOnlySpanToken)
        End Sub

        Friend Function GetReadOnlySpans() As List(Of ReadOnlySpan)
            Return _readOnlySpans
        End Function

        Private Function GetExtentOfReadOnlySpan(span1 As Span) As ReadOnlySpan
            Dim readOnlySpan1 As ReadOnlySpan = Nothing
            Dim currentSnapshot1 As ITextSnapshot = _buffer.CurrentSnapshot

            For Each readOnlySpanToken As IReadOnlyRegionHandle In _readOnlySpanTokens
                Dim span2 As Span = readOnlySpanToken.Span.GetSpan(currentSnapshot1)
                If span2.OverlapsWith(span1) Then
                    readOnlySpan1 = MergeReadOnlySpans(New ReadOnlySpan(currentSnapshot1, span2, readOnlySpanToken.Span.TrackingMode, readOnlySpanToken.EdgeInsertionMode, readOnlySpanToken.EdgeInsertionMode), readOnlySpan1)
                End If
            Next
            Return readOnlySpan1
        End Function

        Private Function MergeReadOnlySpans(first As ReadOnlySpan, second As ReadOnlySpan) As ReadOnlySpan
            If first Is Nothing Then
                Throw New ArgumentNullException("first")
            End If

            If second Is Nothing Then
                Return first
            End If

            Dim currentSnapshot1 As ITextSnapshot = _buffer.CurrentSnapshot
            Dim span1 As Span = first.GetSpan(currentSnapshot1)
            Dim span2 As Span = second.GetSpan(currentSnapshot1)
            Dim start1 As Integer = span1.Start
            Dim [end] As Integer = span1.End
            Dim flag As Boolean = first.StartEdgeInsertionMode = EdgeInsertionMode.Deny
            Dim flag2 As Boolean = first.EndEdgeInsertionMode = EdgeInsertionMode.Deny

            If start1 > span2.Start Then
                start1 = span2.Start
                flag = second.StartEdgeInsertionMode = EdgeInsertionMode.Deny
            ElseIf start1 = span2.Start Then
                flag = flag OrElse second.StartEdgeInsertionMode = EdgeInsertionMode.Deny
            End If

            If [end] < span2.End Then
                [end] = span2.End
                flag2 = second.EndEdgeInsertionMode = EdgeInsertionMode.Deny
            ElseIf [end] = span2.End Then
                flag2 = flag2 OrElse second.EndEdgeInsertionMode = EdgeInsertionMode.Deny
            End If

            Return New ReadOnlySpan(currentSnapshot1, New Span(start1, [end] - start1), first.TrackingMode, If(flag, EdgeInsertionMode.Deny, EdgeInsertionMode.Allow), If(flag2, EdgeInsertionMode.Deny, EdgeInsertionMode.Allow))
        End Function
    End Class
End Namespace
