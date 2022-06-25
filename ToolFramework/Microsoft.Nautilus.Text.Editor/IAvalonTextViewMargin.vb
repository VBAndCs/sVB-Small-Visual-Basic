Imports System.Windows

Namespace Microsoft.Nautilus.Text.Editor
    Public Interface IAvalonTextViewMargin
        Inherits ITextViewMargin
        ReadOnly Property VisualElement As FrameworkElement
    End Interface
End Namespace
