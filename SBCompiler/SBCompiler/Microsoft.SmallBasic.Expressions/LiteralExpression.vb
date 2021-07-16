Imports System
Imports System.Globalization
Imports System.Reflection.Emit

Namespace Microsoft.SmallBasic.Expressions
    <Serializable>
    Public Class LiteralExpression
        Inherits Expression

        Public Sub New()
            MyBase.New()
        End Sub

        Public Sub New(literal As TokenInfo)
            _Literal = literal
        End Sub

        Public Property Literal As TokenInfo

        Public Overrides Sub EmitIL(ByVal scope As CodeGenScope)
            If Literal.Token = Token.StringLiteral Then
                scope.ILGenerator.Emit(OpCodes.Ldstr, Literal.Text.Trim(""""c))
                scope.ILGenerator.EmitCall(OpCodes.Call, scope.TypeInfoBag.StringToPrimitive, Nothing)
            ElseIf Literal.Token = Token.NumericLiteral Then
                scope.ILGenerator.Emit(OpCodes.Ldc_R8, Double.Parse(Literal.Text, CultureInfo.InvariantCulture))
                scope.ILGenerator.EmitCall(OpCodes.Call, scope.TypeInfoBag.NumberToPrimitive, Nothing)
            End If
        End Sub

        Public Overrides Function ToString() As String
            Return Literal.Text
        End Function
    End Class
End Namespace
