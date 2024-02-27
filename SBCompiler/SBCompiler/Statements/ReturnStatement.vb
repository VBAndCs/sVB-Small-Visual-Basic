Imports System.Reflection.Emit
Imports Microsoft.SmallVisualBasic.Completion
Imports Microsoft.SmallVisualBasic.Engine
Imports Microsoft.SmallVisualBasic.Expressions

Namespace Microsoft.SmallVisualBasic.Statements
    Public Class ReturnStatement
        Inherits Statement

        Public ReturnExpression As Expression
        Public Subroutine As SubroutineStatement

        Public Overrides Function GetStatementAt(lineNumber As Integer) As Statement
            If lineNumber < StartToken.Line Then Return Nothing
            If lineNumber <= ReturnExpression?.EndToken.Line Then Return Me
            Return Nothing
        End Function

        Public Overrides Sub AddSymbols(symbolTable As SymbolTable)
            MyBase.AddSymbols(symbolTable)
            If ReturnExpression Is Nothing Then
                If Subroutine?.SubToken.Type = TokenType.Function Then
                    symbolTable.Errors.Add(New [Error](StartToken, "Function must return a value"))
                End If
            Else
                ReturnExpression.Parent = Me
                ReturnExpression.AddSymbols(symbolTable)
            End If

            If Subroutine IsNot Nothing Then
                Subroutine.ReturnStatements.Add(Me)
            End If
        End Sub

        Public Overrides Sub EmitIL(scope As CodeGenScope)
            If scope.ForGlobalHelp Then Return
            ReturnExpression?.EmitIL(scope)
            scope.ILGenerator.Emit(OpCodes.Ret)
        End Sub

        Public Overrides Function ToString() As String
            Return $"{StartToken} {ReturnExpression}" & vbCrLf
        End Function

        Public Overrides Sub PopulateCompletionItems(
                        bag As CompletionBag,
                        line As Integer,
                        column As Integer,
                        globalScope As Boolean
                   )

            If Me.StartToken.IsBefore(line, column) Then
                CompletionHelper.FillExpressionItems(bag)
                CompletionHelper.FillSubroutines(bag, True)
            End If
        End Sub

        Public Overrides Function Execute(runner As ProgramRunner) As statement
            If ReturnExpression IsNot Nothing Then
                runner.Fields($"{Subroutine.Name.LCaseText}.return") = ReturnExpression.Evaluate(runner)
            End If
            Return Me
        End Function
    End Class
End Namespace
