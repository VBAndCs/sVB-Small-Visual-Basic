Imports System
Imports System.Globalization
Imports System.Reflection.Emit

Namespace Microsoft.SmallBasic.Expressions
    <Serializable>
    Public Class NegativeExpression
        Inherits Expression

        Public Property Negation As TokenInfo
        Public Property Expression As Expression

        Public Overrides Sub AddSymbols(ByVal symbolTable As SymbolTable)
            If Expression IsNot Nothing Then
                Expression.Parent = Me.Parent
                Expression.AddSymbols(symbolTable)
            End If
        End Sub

        Public Overrides Sub EmitIL(ByVal scope As CodeGenScope)
            Expression.EmitIL(scope)
            scope.ILGenerator.EmitCall(OpCodes.Call, scope.TypeInfoBag.Negation, Nothing)
        End Sub

        Public Overrides Function ToString() As String
            Return String.Format(CultureInfo.CurrentUICulture, "-{0}", New Object(0) {Expression})
        End Function
    End Class
End Namespace
