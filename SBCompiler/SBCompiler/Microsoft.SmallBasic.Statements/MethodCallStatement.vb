Imports System.Globalization
Imports Microsoft.SmallBasic.Expressions

Namespace Microsoft.SmallBasic.Statements
    Public Class MethodCallStatement
        Inherits Statement

        Public MethodCallExpression As MethodCallExpression

        Public Overrides Sub AddSymbols(ByVal symbolTable As SymbolTable)
            If MethodCallExpression IsNot Nothing Then
                MethodCallExpression.AddSymbols(symbolTable)
            End If
        End Sub

        Public Overrides Sub EmitIL(ByVal scope As CodeGenScope)
            MethodCallExpression.EmitIL(scope)
            Dim methodInfo = MethodCallExpression.GetMethodInfo(scope)
            If methodInfo.ReturnType IsNot GetType(Void) Then
                scope.ILGenerator.Emit(System.Reflection.Emit.OpCodes.Pop)
            End If
        End Sub

        Public Overrides Function ToString() As String
            Return String.Format(CultureInfo.CurrentUICulture, "{0}" & VisualBasic.Constants.vbCrLf, New Object(0) {MethodCallExpression})
        End Function
    End Class
End Namespace
