Imports Microsoft.Nautilus.Text
Imports Microsoft.Nautilus.Text.AdornmentSystem
Imports Microsoft.SmallBasic.Completion

Namespace Microsoft.SmallBasic.LanguageService
    Public NotInheritable Class CompletionAdornment
        Implements IAdornment

        Private _AdornmentProvider As Microsoft.SmallBasic.LanguageService.CompletionAdornmentProvider
        Private completionBagField As CompletionBag
        Private textSpan As ITextSpan

        Public Property AdornmentProvider As CompletionAdornmentProvider
            Get
                Return _AdornmentProvider
            End Get
            Private Set(ByVal value As CompletionAdornmentProvider)
                _AdornmentProvider = value
            End Set
        End Property

        Public ReadOnly Property CompletionBag As CompletionBag
            Get
                Return completionBagField
            End Get
        End Property

        Public ReadOnly Property Span As ITextSpan Implements IAdornment.Span
            Get
                Return textSpan
            End Get
        End Property

        Public Property ReplaceSpan As ITextSpan

        Public Sub New(ByVal provider As CompletionAdornmentProvider, ByVal completionBag As CompletionBag, ByVal adornmentSpan As ITextSpan, ByVal replaceSpan As ITextSpan)
            AdornmentProvider = provider
            completionBagField = completionBag
            textSpan = adornmentSpan
            Me.ReplaceSpan = replaceSpan
        End Sub

        Public Sub Dismiss(ByVal force As Boolean)
            AdornmentProvider.DismissAdornment(force)
        End Sub
    End Class
End Namespace
