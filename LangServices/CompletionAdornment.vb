Imports Microsoft.Nautilus.Text
Imports Microsoft.Nautilus.Text.AdornmentSystem
Imports Microsoft.SmallBasic.Completion

Namespace Microsoft.SmallBasic.LanguageService
    Public NotInheritable Class CompletionAdornment
        Implements IAdornment

        Private _completionBag As CompletionBag
        Private textSpan As ITextSpan

        Public Property AdornmentProvider As CompletionAdornmentProvider

        Public ReadOnly Property CompletionBag As CompletionBag
            Get
                Return _completionBag
            End Get
        End Property

        Public ReadOnly Property Span As ITextSpan Implements IAdornment.Span
            Get
                Return textSpan
            End Get
        End Property

        Public Property ReplaceSpan As ITextSpan

        Public Sub New(
                      provider As CompletionAdornmentProvider,
                       completionBag As CompletionBag,
                       adornmentSpan As ITextSpan,
                       replaceSpan As ITextSpan)

            _AdornmentProvider = provider
            _completionBag = completionBag
            textSpan = adornmentSpan
            _ReplaceSpan = replaceSpan
        End Sub

        Public Sub Dismiss(force As Boolean)
            AdornmentProvider.DismissAdornment(force)
        End Sub
    End Class
End Namespace
