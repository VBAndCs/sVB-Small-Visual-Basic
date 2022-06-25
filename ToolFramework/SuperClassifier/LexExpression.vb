Imports Microsoft.Nautilus.Text

Namespace SuperClassifier
    Friend MustInherit Class LexExpression
        Friend ReadOnly ExpressionKind As LexExpressionKind
        Friend ReadOnly Classification As String

        Friend Sub New(exprKind As LexExpressionKind, classification As String)
            ExpressionKind = exprKind
            Me.Classification = classification
        End Sub

        Friend MustOverride Function GetTokenStartingAt(textSnapshot As ITextSnapshot, start As Integer) As Token
    End Class
End Namespace
