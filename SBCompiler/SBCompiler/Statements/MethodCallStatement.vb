Imports System.Globalization
Imports Microsoft.SmallBasic.Expressions

Namespace Microsoft.SmallBasic.Statements
    Public Class MethodCallStatement
        Inherits Statement

        Public MethodCallExpression As MethodCallExpression
        Public Shared DoNotPopReturnValue As Boolean

        Public Overrides Function GetStatement(lineNumber) As Statement
            If lineNumber < StartToken.Line Then Return Nothing
            If lineNumber <= MethodCallExpression?.EndToken.Line Then Return Me
            Return Nothing
        End Function

        Public Overrides Sub AddSymbols(ByVal symbolTable As SymbolTable)
            MyBase.AddSymbols(symbolTable)
            If MethodCallExpression IsNot Nothing Then
                MethodCallExpression.Parent = Me
                MethodCallExpression.AddSymbols(symbolTable)
            End If
        End Sub

        Public Overrides Sub EmitIL(ByVal scope As CodeGenScope)
            MethodCallExpression.EmitIL(scope)
            Dim methodInfo = MethodCallExpression.GetMethodInfo(scope)
            If Not DoNotPopReturnValue AndAlso methodInfo.ReturnType IsNot GetType(Void) Then
                scope.ILGenerator.Emit(System.Reflection.Emit.OpCodes.Pop)
            End If
        End Sub

        Public Overrides Function ToString() As String
            Return String.Format(CultureInfo.CurrentUICulture, "{0}" & vbCrLf, New Object(0) {MethodCallExpression})
        End Function
    End Class
End Namespace
