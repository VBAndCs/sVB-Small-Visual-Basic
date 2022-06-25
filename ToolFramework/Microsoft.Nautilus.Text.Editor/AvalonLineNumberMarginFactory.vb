Imports System.ComponentModel.Composition
Imports Microsoft.Nautilus.Core
Imports Microsoft.Nautilus.Text.Operations

Namespace Microsoft.Nautilus.Text.Editor

    <AvalonTextViewMarginPlacement(MarginPlacement.Left)>
    <Name("Avalon Line Number Margin")>
    <ComponentOptions(ComponentDiscoveryMode:=ComponentDiscoveryMode.Always)>
    Public NotInheritable Class AvalonLineNumberMarginFactory
        Inherits AvalonTextViewMarginFactory

        <Import>
        Public Property EditorOperationsProvider As IEditorOperationsProvider

        Public Overrides Function CreateAvalonTextViewMargin(textView As IAvalonTextView) As IAvalonTextViewMargin
            Return New AvalonLineNumberMargin(textView, _editorOperationsProvider.GetEditorOperations(textView))
        End Function

    End Class
End Namespace
