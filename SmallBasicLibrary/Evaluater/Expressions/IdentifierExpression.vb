Imports System
Imports Microsoft.SmallVisualBasic.Library

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
                    Return Library.Math.E
                Case "pi"
                    Return Library.Math.Pi
                Case Else
                    Evaluator.Errors.Add(New [Error](
                           Identifier,
                            $"Unkown Identifier {Identifier.Text}"
                    ))
                    Return New Primitive("")
            End Select
        End Function
    End Class
End Namespace
