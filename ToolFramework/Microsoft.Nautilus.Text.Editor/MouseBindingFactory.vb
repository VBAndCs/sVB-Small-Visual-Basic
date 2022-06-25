Imports System.ComponentModel.Composition

Namespace Microsoft.Nautilus.Text.Editor
    <Export>
    Public MustInherit Class MouseBindingFactory
        Public MustOverride Function GetAssociatedBinding(avalonTextViewHost As IAvalonTextViewHost) As IMouseBinding
    End Class
End Namespace
