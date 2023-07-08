Imports System

Namespace Evaluator.Expressions
    <Serializable>
    Friend Class IdentifierExpression
        Inherits Expression

        Public Identifier As Token

        Public Overrides Function ToString() As String
            Return Identifier.Text
        End Function

        Public Overrides Function Evaluate(x As Double) As Library.Primitive
            Select Case Identifier.LCaseText
                Case "x"
                    Return x
                Case "e"
                    Return Math.E
                Case "pi"
                    Return Math.PI
                Case Else
                    Evaluator.Errors.Add(New [Error](
                           Identifier,
                            $"Unkown Identifier {Identifier.Text}"
                    ))
                    Return ""
            End Select
        End Function
    End Class
End Namespace
