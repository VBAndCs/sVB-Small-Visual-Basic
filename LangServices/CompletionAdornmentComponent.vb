Imports System.ComponentModel.Composition
Imports Microsoft.Nautilus.Core.Undo
Imports Microsoft.Nautilus.Text
Imports Microsoft.Nautilus.Text.AdornmentSystem
Imports Microsoft.Nautilus.Text.Editor
Imports Microsoft.Nautilus.Text.Operations

Namespace Microsoft.SmallBasic.LanguageService
    Public NotInheritable Class CompletionAdornmentComponent
        <Import>
        Public Property EditorOperationsProvider As IEditorOperationsProvider
        <Import>
        Public Property UndoHistoryRegistry As IUndoHistoryRegistry

        <Export(GetType(AdornmentProviderFactory))>
        <ContentType("text.smallbasic")>
        Public Function CreateAdornmentProvider(ByVal textView As ITextView) As IAdornmentProvider
            Dim completionAdornmentProvider As CompletionAdornmentProvider = New CompletionAdornmentProvider(textView, EditorOperationsProvider, UndoHistoryRegistry)
            textView.Properties.AddProperty(GetType(CompletionAdornmentProvider), completionAdornmentProvider)
            Return completionAdornmentProvider
        End Function

        <AdornmentDiscriminator(GetType(CompletionAdornment))>
        <Export(GetType(AdornmentSurfaceFactory))>
        Public Function CreateCompletionItemsSurface(ByVal textView As IAvalonTextView) As IAdornmentSurface
            Dim completionAdornmentSurface As CompletionAdornmentSurface = New CompletionAdornmentSurface(textView)
            textView.Properties.AddProperty(GetType(CompletionAdornmentSurface), completionAdornmentSurface)
            Return completionAdornmentSurface
        End Function
    End Class
End Namespace
