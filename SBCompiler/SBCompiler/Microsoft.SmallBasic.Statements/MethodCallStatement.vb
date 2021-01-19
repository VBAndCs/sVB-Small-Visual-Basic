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
                        ''' Cannot convert IfStatementSyntax, System.InvalidCastException: Unable to cast object of type 'Microsoft.CodeAnalysis.VisualBasic.Syntax.EmptyStatementSyntax' to type 'Microsoft.CodeAnalysis.VisualBasic.Syntax.ExpressionSyntax'.
'''    at ICSharpCode.CodeConverter.VB.MethodBodyExecutableStatementVisitor.VisitIfStatement(IfStatementSyntax node)
'''    at Microsoft.CodeAnalysis.CSharp.Syntax.IfStatementSyntax.Accept[TResult](CSharpSyntaxVisitor`1 visitor)
'''    at Microsoft.CodeAnalysis.CSharp.CSharpSyntaxVisitor`1.Visit(SyntaxNode node)
'''    at ICSharpCode.CodeConverter.VB.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node)
''' 
''' Input:
''' 			if (methodInfo.ReturnType != typeof(void))
''' 			{
''' 				scope.ILGenerator.Emit(System.Reflection.Emit.OpCodes.Pop);
''' 			}
''' 
''' 
        End Sub

        Public Overrides Function ToString() As String
            Return String.Format(CultureInfo.CurrentUICulture, "{0}" & VisualBasic.Constants.vbCrLf, New Object(0) {MethodCallExpression})
        End Function
    End Class
End Namespace
