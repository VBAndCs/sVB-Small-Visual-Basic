
Namespace Evaluator.Expressions
    <Serializable>
    Friend Class BinaryExpression
        Inherits Expression

        Public Property LeftHandSide As Expression
        Public Property [Operator] As Token
        Public Property RightHandSide As Expression


        Public Overrides Function ToString() As String
            Return $"({LeftHandSide} {[Operator].Text} {RightHandSide})"
        End Function

        Public Overrides Function Evaluate(x As Double) As Library.Primitive
            Select Case [Operator].Type
                Case TokenType.Addition
                    Return LeftHandSide.Evaluate(x) + RightHandSide.Evaluate(x)
                Case TokenType.Subtraction
                    Return LeftHandSide.Evaluate(x) - RightHandSide.Evaluate(x)
                Case TokenType.Multiplication
                    Return LeftHandSide.Evaluate(x) * RightHandSide.Evaluate(x)
                Case TokenType.Division
                    Return LeftHandSide.Evaluate(x) / RightHandSide.Evaluate(x)
                Case TokenType.Mod
                    Return LeftHandSide.Evaluate(x).AsDecimal Mod RightHandSide.Evaluate(x).AsDecimal
                Case TokenType.Power
                    Return LeftHandSide.Evaluate(x).AsDecimal ^ RightHandSide.Evaluate(x).AsDecimal
            End Select

            Evaluator.Errors.Add(New [Error](
                     [Operator],
                     $"Unrecognized operator {[Operator].Text}"
           ))
            Return ""
        End Function
    End Class
End Namespace
