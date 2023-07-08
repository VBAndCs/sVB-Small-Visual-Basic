
Namespace Evaluator.Expressions
    <Serializable>
    Friend Class LiteralExpression
        Inherits Expression

        Public Sub New()
            MyBase.New()
        End Sub

        Public Sub New(literal As Token)
            _Literal = literal
        End Sub

        Public Property Literal As Token

        Public Shared ReadOnly Zero As New LiteralExpression(
            New Token() With {
                    .Type = TokenType.NumericLiteral,
                    .Text = 0
            }
        )

        Public Overrides Function ToString() As String
            Return Literal.Text
        End Function

        Public Overrides Function Evaluate(x As Double) As Library.Primitive
            Return New Library.Primitive(CDbl(Literal.Text))
        End Function
    End Class

End Namespace
