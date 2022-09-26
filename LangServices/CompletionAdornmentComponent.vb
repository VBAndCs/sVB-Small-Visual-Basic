Imports System.ComponentModel.Composition
Imports Microsoft.Nautilus.Core.Undo
Imports Microsoft.Nautilus.Text
Imports Microsoft.Nautilus.Text.AdornmentSystem
Imports Microsoft.Nautilus.Text.Editor
Imports Microsoft.Nautilus.Text.Operations

Namespace Microsoft.SmallVisualBasic.LanguageService
    Public NotInheritable Class CompletionAdornmentComponent
        <Import>
        Public Property EditorOperationsProvider As IEditorOperationsProvider
        <Import>
        Public Property UndoHistoryRegistry As IUndoHistoryRegistry

        <Export(GetType(AdornmentProviderFactory))>
        <ContentType("text.smallbasic")>
        Public Function CreateAdornmentProvider(textView As ITextView) As IAdornmentProvider
            Dim provider As New CompletionProvider(textView, EditorOperationsProvider, UndoHistoryRegistry)
            textView.Properties.AddProperty(GetType(CompletionProvider), provider)
            Return provider
        End Function

        <AdornmentDiscriminator(GetType(CompletionAdornment))>
        <Export(GetType(AdornmentSurfaceFactory))>
        Public Function CreateCompletionItemsSurface(textView As IAvalonTextView) As IAdornmentSurface
            Dim completionAdornmentSurface As CompletionAdornmentSurface = New CompletionAdornmentSurface(textView)
            textView.Properties.AddProperty(GetType(CompletionAdornmentSurface), completionAdornmentSurface)
            Return completionAdornmentSurface
        End Function
    End Class
End Namespace
