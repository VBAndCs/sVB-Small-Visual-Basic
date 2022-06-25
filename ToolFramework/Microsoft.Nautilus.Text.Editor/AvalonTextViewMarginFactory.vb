Imports System.ComponentModel.Composition

Namespace Microsoft.Nautilus.Text.Editor
    <Export>
    Public MustInherit Class AvalonTextViewMarginFactory
        Public MustOverride Function CreateAvalonTextViewMargin(avalonTextView As IAvalonTextView) As IAvalonTextViewMargin
    End Class
End Namespace
