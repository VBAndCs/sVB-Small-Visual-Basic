Imports System

Namespace Evaluator.Expressions
    <Serializable>
    Friend MustInherit Class Expression
        Public Property StartToken As Token
        Public Property EndToken As Token
        Public Property Precedence As Integer

        Public MustOverride Function Evaluate(x As Double) As Library.Primitive
    End Class
End Namespace
