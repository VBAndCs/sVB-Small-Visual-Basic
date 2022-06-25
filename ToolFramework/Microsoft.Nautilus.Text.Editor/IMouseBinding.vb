Imports System.Windows.Input

Namespace Microsoft.Nautilus.Text.Editor
    Public Interface IMouseBinding
        Sub PreprocessMouseLeftButtonDown(e As MouseButtonEventArgs)

        Sub PostprocessMouseLeftButtonDown(e As MouseButtonEventArgs)

        Sub PreprocessMouseRightButtonDown(e As MouseButtonEventArgs)

        Sub PostprocessMouseRightButtonDown(e As MouseButtonEventArgs)

        Sub PreprocessMouseLeftButtonUp(e As MouseButtonEventArgs)

        Sub PostprocessMouseLeftButtonUp(e As MouseButtonEventArgs)

        Sub PreprocessMouseRightButtonUp(e As MouseButtonEventArgs)

        Sub PostprocessMouseRightButtonUp(e As MouseButtonEventArgs)

        Sub PreprocessMouseMove(e As MouseEventArgs)

        Sub PostprocessMouseMove(e As MouseEventArgs)

        Sub PreprocessMouseWheel(e As MouseWheelEventArgs)

        Sub PostprocessMouseWheel(e As MouseWheelEventArgs)

        Sub PreprocessMouseEnter(e As MouseEventArgs)

        Sub PostprocessMouseEnter(e As MouseEventArgs)

        Sub PreprocessMouseLeave(e As MouseEventArgs)

        Sub PostprocessMouseLeave(e As MouseEventArgs)
    End Interface
End Namespace
