Imports System.Windows
Imports System.Windows.Controls.Primitives
Imports System.Windows.Media

Namespace Microsoft.Nautilus.Text.Editor
    Public Interface IAvalonTextView
        Inherits ITextView, IPropertyOwner
        ReadOnly Property VisualElement As FrameworkElement

        Property Background As Brush

        ReadOnly Property SpanGeometry As ISpanGeometry

        Sub Invalidate()
        Sub OnScrollChanged(e As ScrollEventArgs)

        Event ScrollChaged(senmder As Object, e As ScrollEventArgs)
    End Interface
End Namespace
