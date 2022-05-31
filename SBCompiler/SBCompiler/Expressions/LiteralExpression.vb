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

        Public Overrides Sub AddSymbols(symbolTable As SymbolTable)
            MyBase.AddSymbols(symbolTable)
            _Literal.Parent = Me.Parent
        End Sub

        Public Overrides Sub EmitIL(ByVal scope As CodeGenScope)
            Select Case Literal.Token
                Case Token.StringLiteral
                    scope.ILGenerator.Emit(OpCodes.Ldstr, Literal.Text.Trim(""""c))
                    scope.ILGenerator.EmitCall(OpCodes.Call, scope.TypeInfoBag.StringToPrimitive, Nothing)

                Case Token.True
                    scope.ILGenerator.Emit(OpCodes.Ldstr, "True")
                    scope.ILGenerator.EmitCall(OpCodes.Call, scope.TypeInfoBag.StringToPrimitive, Nothing)

                Case Token.False
                    scope.ILGenerator.Emit(OpCodes.Ldstr, "False")
                    scope.ILGenerator.EmitCall(OpCodes.Call, scope.TypeInfoBag.StringToPrimitive, Nothing)

                Case Token.NumericLiteral
                    scope.ILGenerator.Emit(OpCodes.Ldc_R8, Double.Parse(Literal.Text, CultureInfo.InvariantCulture))
                    scope.ILGenerator.EmitCall(OpCodes.Call, scope.TypeInfoBag.NumberToPrimitive, Nothing)
            End Select
        End Sub

        Public Overrides Function ToString() As String
            Return Literal.Text
        End Function
    End Class

End Namespace
