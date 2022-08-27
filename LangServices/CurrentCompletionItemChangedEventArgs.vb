Imports System

Namespace Microsoft.SmallBasic.LanguageService
    Public Class CurrentCompletionItemChangedEventArgs
        Inherits EventArgs

        Private _currentCompletionItem As CompletionItemWrapper

        Public ReadOnly Property CurrentCompletionItem As CompletionItemWrapper
            Get
                Return _currentCompletionItem
            End Get
        End Property

        Public Sub New(completionItem As CompletionItemWrapper)
            _currentCompletionItem = completionItem
        End Sub
    End Class
End Namespace
