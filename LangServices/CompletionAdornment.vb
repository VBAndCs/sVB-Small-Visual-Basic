Imports Microsoft.Nautilus.Text
Imports Microsoft.Nautilus.Text.AdornmentSystem
Imports Microsoft.SmallBasic.Completion

Namespace Microsoft.SmallBasic.LanguageService
    Public NotInheritable Class CompletionAdornment
        Implements IAdornment

        Public Property AdornmentProvider As CompletionProvider

        Public ReadOnly Property CompletionBag As CompletionBag

        Public ReadOnly Property Span As ITextSpan Implements IAdornment.Span


        Public Property ReplaceSpan As ITextSpan

        Public Sub New(
                      provider As CompletionProvider,
                       completionBag As CompletionBag,
                       adornmentSpan As ITextSpan,
                       replaceSpan As ITextSpan)

            _AdornmentProvider = provider
            _CompletionBag = completionBag
            _Span = adornmentSpan
            _ReplaceSpan = replaceSpan
        End Sub

        Public Sub Dismiss(force As Boolean)
            AdornmentProvider.DismissAdornment(force)
        End Sub
    End Class
End Namespace
