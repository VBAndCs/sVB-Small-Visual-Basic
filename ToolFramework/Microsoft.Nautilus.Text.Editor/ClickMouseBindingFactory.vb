Imports System.ComponentModel.Composition
Imports Microsoft.Nautilus.Text.Operations

Namespace Microsoft.Nautilus.Text.Editor
    <ComponentOptions(ComponentDiscoveryMode:=ComponentDiscoveryMode.Always)>
    Public NotInheritable Class ClickMouseBindingFactory
        Inherits MouseBindingFactory

        <Import>
        Public Property EditorOperationsProvider As IEditorOperationsProvider

        <Import>
        Public Property TextStructureNavigatorFactory As ITextStructureNavigatorFactory

        Public Overrides Function GetAssociatedBinding(avalonTextViewHost As IAvalonTextViewHost) As IMouseBinding
            If avalonTextViewHost Is Nothing Then
                Throw New ArgumentNullException("avalonTextViewHost")
            End If

            Return New ClickMouseBinding(avalonTextViewHost.TextView, _editorOperationsProvider, _textStructureNavigatorFactory)
        End Function
    End Class
End Namespace
