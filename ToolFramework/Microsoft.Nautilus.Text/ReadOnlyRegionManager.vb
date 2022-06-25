Imports System.Collections.Generic

Namespace Microsoft.Nautilus.Text
    Public MustInherit Class ReadOnlyRegionManager

        Public Event ReadOnlyRegionChanged As EventHandler(Of ReadOnlyRegionChangedEventArgs)

        Protected Sub OnReadOnlyRegionCreated(handle As IReadOnlyRegionHandle)
            RaiseEvent ReadOnlyRegionChanged(Me, New ReadOnlyRegionChangedEventArgs(handle, ReadOnlyRegionChange.Created))
        End Sub

        Protected Sub OnReadOnlyRegionRemoved(handle As IReadOnlyRegionHandle)
            RaiseEvent ReadOnlyRegionChanged(Me, New ReadOnlyRegionChangedEventArgs(handle, ReadOnlyRegionChange.Removed))
        End Sub

        Public Function CreateReadOnlyRegion(position As Integer, length As Integer) As IReadOnlyRegionHandle
            Return CreateReadOnlyRegion(New Span(position, length))
        End Function

        Public Function CreateReadOnlyRegion(span1 As Span) As IReadOnlyRegionHandle
            Return CreateReadOnlyRegion(span1, SpanTrackingMode.EdgeExclusive)
        End Function

        Public Function CreateReadOnlyRegion(position As Integer, length As Integer, trackingMode As SpanTrackingMode) As IReadOnlyRegionHandle
            Return CreateReadOnlyRegion(New Span(position, length), trackingMode)
        End Function

        Public Function CreateReadOnlyRegion(span1 As Span, trackingMode As SpanTrackingMode) As IReadOnlyRegionHandle
            Return CreateReadOnlyRegion(span1, trackingMode, EdgeInsertionMode.Allow)
        End Function

        Public Function CreateReadOnlyRegion(position As Integer, length As Integer, trackingMode As SpanTrackingMode, edgeInsertionMode1 As EdgeInsertionMode) As IReadOnlyRegionHandle
            Return CreateReadOnlyRegion(New Span(position, length), trackingMode, edgeInsertionMode1)
        End Function

        Public MustOverride Function CreateReadOnlyRegion(span1 As Span, trackingMode As SpanTrackingMode, edgeInsertionMode1 As EdgeInsertionMode) As IReadOnlyRegionHandle

        Public Function GetReadOnlyExtents(position As Integer, length As Integer) As IList(Of Span)
            Return GetReadOnlyExtents(New Span(position, length))
        End Function

        Public MustOverride Function GetReadOnlyExtents(span1 As Span) As IList(Of Span)
        Protected MustOverride Function CreateReadOnlyRegion(snapshot As ITextSnapshot, span1 As Span, trackingMode As SpanTrackingMode, edgeInsertionMode1 As EdgeInsertionMode) As IReadOnlyRegionHandle
        Protected MustOverride Sub RemoveReadOnlyRegion(readOnlyRegionToken As IReadOnlyRegionHandle)

        Friend Sub RemoveReadOnlyRegionHandle(readOnlyRegionToken As IReadOnlyRegionHandle)
            RemoveReadOnlyRegion(readOnlyRegionToken)
        End Sub
    End Class
End Namespace
