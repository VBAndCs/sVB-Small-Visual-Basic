Imports System
Imports System.Globalization
Imports System.Reflection.Emit

Namespace Microsoft.SmallVisualBasic.Expressions
    <Serializable>
    Public Class NegativeExpression
        Inherits Expression

        Public Property Negation As Token
        Public Property Expression As Expression

        Public Overrides Sub AddSymbols(symbolTable As SymbolTable)
            MyBase.AddSymbols(symbolTable)
            _Negation.Parent = Me.Parent

            If Expression IsNot Nothing Then
                Expression.Parent = Me.Parent
                Expression.AddSymbols(symbolTable)
            End If
        End Sub

        Public Overrides Sub EmitIL(scope As CodeGenScope)
            If Expression Is Nothing Then
                LiteralExpression.Zero.EmitIL(scope)
            Else
                Expression.EmitIL(scope)
            End If

            scope.ILGenerator.EmitCall(OpCodes.Call, scope.TypeInfoBag.Negation, Nothing)
        End Sub

        Public Overrides Function ToString() As String
            Return String.Format(CultureInfo.CurrentUICulture, "-{0}", New Object(0) {Expression})
        End Function
    End Class
End Namespace
