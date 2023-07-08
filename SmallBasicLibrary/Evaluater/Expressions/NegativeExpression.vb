
Namespace Evaluator.Expressions
    <Serializable>
    Friend Class NegativeExpression
        Inherits Expression

        Public Property Negation As Token
        Public Property Expression As Expression

        Public Overrides Function ToString() As String
            Return $"-{Expression}"
        End Function

        Public Overrides Function Evaluate(x As Double) As Library.Primitive
            Return -Expression.Evaluate(x)
        End Function
    End Class
End Namespace
