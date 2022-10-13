Imports Microsoft.Nautilus.Text
Imports Microsoft.Nautilus.Text.AdornmentSystem
Imports Microsoft.SmallVisualBasic.Completion

Namespace Microsoft.SmallVisualBasic.LanguageService
    Public NotInheritable Class CompletionAdornment
        Implements IAdornment

        Public Property AdornmentProvider As CompletionProvider
        Public Property AdornmentSurface As CompletionSurface

        Public ReadOnly Property CompletionBag As CompletionBag

        Public ReadOnly Property Span As ITextSpan Implements IAdornment.Span


        Public Property ReplaceSpan As ITextSpan

        Public ReadOnly Property IsVisible
            Get
                Return _AdornmentSurface.IsAdornmentVisible
            End Get
        End Property

        Public Sub ModifySpans(adornmentSpan As ITextSpan, replaceSpan As ITextSpan)
            _Span = adornmentSpan
            _ReplaceSpan = replaceSpan
        End Sub

        Public Sub New(
                      provider As CompletionProvider,
                       completionBag As CompletionBag,
                       adornmentSpan As ITextSpan,
                       replaceSpan As ITextSpan
                   )

            _AdornmentProvider = provider
            _AdornmentSurface = provider.textView.Properties.GetProperty(Of CompletionSurface)()

            _CompletionBag = completionBag
            _Span = adornmentSpan
            _ReplaceSpan = replaceSpan
        End Sub

        Public Sub Dismiss(force As Boolean)
            AdornmentProvider.DismissAdornment(force)
        End Sub
    End Class
End Namespace
