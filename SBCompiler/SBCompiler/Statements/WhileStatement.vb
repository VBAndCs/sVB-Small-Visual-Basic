Imports System.Reflection.Emit
Imports System.Text
Imports Microsoft.SmallVisualBasic.Completion
Imports Microsoft.SmallVisualBasic.Expressions

Namespace Microsoft.SmallVisualBasic.Statements
    Public Class WhileStatement
        Inherits LoopStatement

        Public WhileToken As Token
        Public Condition As Expression

        Public Overrides Function GetStatementAt(lineNumber As Integer) As Statement
            If lineNumber < StartToken.Line Then Return Nothing
            If lineNumber > EndLoopToken.Line Then Return Nothing

            If lineNumber <= Condition?.EndToken.Line Then Return Me

            For Each statment In Body
                Dim st = statment.GetStatementAt(lineNumber)
                If st IsNot Nothing Then Return st
            Next

            If lineNumber <= EndLoopToken.Line Then Return Me

            Return Nothing
        End Function

        Public Overrides Function GetKeywords() As LegalTokens
            Dim spans As New LegalTokens
            spans.Add(WhileToken)
            For Each statement In JumpLoopStatements
                spans.Add(statement.StartToken)
            Next
            spans.Add(EndLoopToken)
            Return spans
        End Function

        Public Overrides Sub AddSymbols(symbolTable As SymbolTable)
            MyBase.AddSymbols(symbolTable)
            WhileToken.Parent = Me
            EndLoopToken.Parent = Me

            If Condition IsNot Nothing Then
                Condition.Parent = Me
                Condition.AddSymbols(symbolTable)
            End If

            For Each item In Body
                item.Parent = Me
                item.AddSymbols(symbolTable)
            Next
        End Sub

        Public Overrides Sub PrepareForEmit(scope As CodeGenScope)
            If scope.ForGlobalHelp Then Return

            For Each item In Body
                item.PrepareForEmit(scope)
            Next
        End Sub

        Public Overrides Sub EmitIL(scope As CodeGenScope)
            If scope.ForGlobalHelp Then Return

            ExitLabel = scope.ILGenerator.DefineLabel()
            ContinueLabel = scope.ILGenerator.DefineLabel()

            scope.ILGenerator.MarkLabel(ContinueLabel)
            Condition.EmitIL(scope)
            scope.ILGenerator.EmitCall(OpCodes.Call, scope.TypeInfoBag.PrimitiveToBoolean, Nothing)
            scope.ILGenerator.Emit(OpCodes.Brfalse, ExitLabel)

            For Each item In Body
                item.EmitIL(scope)
            Next

            scope.ILGenerator.Emit(OpCodes.Br, ContinueLabel)
            scope.ILGenerator.MarkLabel(ExitLabel)
        End Sub

        Public Overrides Sub PopulateCompletionItems(
                            completionBag As CompletionBag,
                            line As Integer,
                            column As Integer,
                            globalScope As Boolean
                   )

            If StartToken.Line = line Then
                If column <= WhileToken.EndColumn Then
                    CompletionHelper.FillAllGlobalItems(completionBag, globalScope)
                Else
                    CompletionHelper.FillSubroutines(completionBag, functionsOnly:=True)
                    CompletionHelper.FillLogicalExpressionItems(completionBag)
                End If

            ElseIf Not EndLoopToken.IsIllegal AndAlso line = EndLoopToken.Line Then
                completionBag.CompletionItems.Clear()

            Else
                Dim statementContaining = GetStatementContaining(Body, line)

                If statementContaining IsNot Nothing Then
                    If EndLoopToken.IsIllegal Then CompletionHelper.FillKeywords(completionBag, {TokenType.Wend})
                    statementContaining.PopulateCompletionItems(completionBag, line, column, globalScope:=False)
                End If
            End If
        End Sub

        Public Overrides Function ToString() As String
            Dim sb As New StringBuilder()
            sb.AppendLine($"{WhileToken.Text} {Condition}")

            For Each st In Body
                sb.Append("  ")
                sb.Append(st.ToString())
            Next

            sb.AppendLine(EndLoopToken.Text)
            Return sb.ToString()
        End Function

    End Class
End Namespace
