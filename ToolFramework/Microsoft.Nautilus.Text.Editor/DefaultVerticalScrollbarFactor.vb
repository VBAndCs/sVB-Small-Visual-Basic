Imports System.ComponentModel.Composition
Imports Microsoft.Nautilus.Core

Namespace Microsoft.Nautilus.Text.Editor

    <AvalonTextViewMarginPlacement(MarginPlacement.Right)>
    <ComponentOptions(ComponentDiscoveryMode:=ComponentDiscoveryMode.Always)>
    <Name("Avalon Vertical Scrollbar")>
    Public NotInheritable Class DefaultVerticalScrollbarFactory
        Inherits AvalonTextViewMarginFactory

        Public Overrides Function CreateAvalonTextViewMargin(textView As IAvalonTextView) As IAvalonTextViewMargin
            Return New DefaultVerticalScrollbar(textView)
        End Function
    End Class
End Namespace
