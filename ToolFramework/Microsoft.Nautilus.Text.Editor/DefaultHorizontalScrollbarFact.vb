Imports System.ComponentModel.Composition
Imports Microsoft.Nautilus.Core

Namespace Microsoft.Nautilus.Text.Editor

    <Name("Avalon Horizontal Scrollbar")>
    <AvalonTextViewMarginPlacement(MarginPlacement.Bottom)>
    <ComponentOptions(ComponentDiscoveryMode:=ComponentDiscoveryMode.Always)>
    Public NotInheritable Class DefaultHorizontalScrollbarFactory
        Inherits AvalonTextViewMarginFactory
        Public Overrides Function CreateAvalonTextViewMargin(textView As IAvalonTextView) As IAvalonTextViewMargin
            Return New DefaultHorizontalScrollbar(textView)
        End Function
    End Class
End Namespace
