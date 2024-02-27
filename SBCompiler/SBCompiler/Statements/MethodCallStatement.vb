Imports Microsoft.SmallVisualBasic.Completion
Imports Microsoft.SmallVisualBasic.Engine
Imports Microsoft.SmallVisualBasic.Expressions

Namespace Microsoft.SmallVisualBasic.Statements
    Public Class MethodCallStatement
        Inherits Statement

        Public MethodCallExpression As MethodCallExpression

        Public Overrides Function GetStatementAt(lineNumber As Integer) As Statement
            If lineNumber < StartToken.Line Then Return Nothing
            If lineNumber <= MethodCallExpression?.EndToken.Line Then Return Me
            Return Nothing
        End Function

        Public Overrides Sub AddSymbols(symbolTable As SymbolTable)
            MyBase.AddSymbols(symbolTable)
            If MethodCallExpression IsNot Nothing Then
                MethodCallExpression.Parent = Me
                MethodCallExpression.AddSymbols(symbolTable)
            End If
        End Sub

        Public Overrides Sub EmitIL(scope As CodeGenScope)
            If scope.ForGlobalHelp Then Return

            MethodCallExpression.EmitIL(scope)
            Dim methodInfo = MethodCallExpression.GetMethodInfo(scope)
            If methodInfo.ReturnType IsNot GetType(Void) Then
                scope.ILGenerator.Emit(System.Reflection.Emit.OpCodes.Pop)
            End If
        End Sub

        Public Overrides Sub PopulateCompletionItems(
                        bag As CompletionBag,
                        line As Integer,
                        column As Integer,
                        globalScope As Boolean
                   )

            Dim args = MethodCallExpression.Arguments
            If args.Count > 0 AndAlso line >= args(0).StartToken.Line Then
                CompletionHelper.FillExpressionItems(bag)
                CompletionHelper.FillSubroutines(bag, True)
            Else
                CompletionHelper.FillAllGlobalItems(bag, globalScope)
            End If
        End Sub

        Public Overrides Function ToString() As String
            Return MethodCallExpression.ToString() & vbCrLf
        End Function

        Public Overrides Function Execute(runner As ProgramRunner) As statement
            MethodCallExpression.Evaluate(runner)
            Return Nothing
        End Function
    End Class
End Namespace
