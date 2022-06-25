Namespace Microsoft.Nautilus.Text.Editor
    Public Interface IViewScroller
        Sub ScrollViewportVerticallyByPixels(distanceToScroll As Double)

        Sub ScrollViewportVerticallyByLine(direction As ScrollDirection)

        Sub ScrollViewportVerticallyByPage(direction As ScrollDirection)

        Sub ScrollViewportHorizontallyByPixels(distanceToScroll As Double)

        Function EnsureSpanVisible(span1 As Span, horizontalPadding As Double, verticalPadding As Double) As Boolean
    End Interface
End Namespace
