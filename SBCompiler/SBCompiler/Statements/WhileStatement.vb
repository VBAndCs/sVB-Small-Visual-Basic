Imports System.Collections.Generic
Imports System.Globalization
Imports System.Reflection.Emit
Imports System.Text
Imports Microsoft.SmallBasic.Completion
Imports Microsoft.SmallBasic.Expressions

Namespace Microsoft.SmallBasic.Statements
    Public Class WhileStatement
        Inherits LoopStatement

        Public WhileToken As TokenInfo
        Public Condition As Expression

        Public Overrides Function GetStatement(lineNumber) As Statement
            If lineNumber < StartToken.Line Then Return Nothing
            If lineNumber > EndLoopToken.Line Then Return Nothing

            If lineNumber <= Condition?.EndToken.Line Then Return Me

            For Each statment In Body
                Dim st = statment.GetStatement(lineNumber)
                If st IsNot Nothing Then Return st
            Next

            If lineNumber <= EndLoopToken.Line Then Return Me

            Return Nothing
        End Function

        Public Overrides Function GetKeywords() As LegalTokens
            Dim spans As New LegalTokens
            spans.Add(WhileToken)
            spans.Add(EndLoopToken)
            Return spans
        End Function

        Public Overrides Sub AddSymbols(ByVal symbolTable As SymbolTable)
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

        Public Overrides Sub PrepareForEmit(ByVal scope As CodeGenScope)
            For Each item In Body
                item.PrepareForEmit(scope)
            Next
        End Sub

        Public Overrides Sub EmitIL(ByVal scope As CodeGenScope)
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

        Public Overrides Sub PopulateCompletionItems(ByVal completionBag As CompletionBag, ByVal line As Integer, ByVal column As Integer, ByVal globalScope As Boolean)
            If StartToken.Line = line Then
                If column <= WhileToken.EndColumn Then
                    CompletionHelper.FillAllGlobalItems(completionBag, globalScope)
                Else
                    CompletionHelper.FillLogicalExpressionItems(completionBag)
                End If

                Return
            End If

            Dim statementContaining = GetStatementContaining(Body, line)

            If statementContaining IsNot Nothing Then
                CompletionHelper.FillKeywords(completionBag, Token.EndWhile, Token.Wend)
                statementContaining.PopulateCompletionItems(completionBag, line, column, globalScope:=False)
            End If
        End Sub

        Public Overrides Function ToString() As String
            Dim stringBuilder As StringBuilder = New StringBuilder()
            stringBuilder.AppendFormat(CultureInfo.CurrentUICulture, "{0} {1}", New Object(1) {WhileToken.Text, Condition})
            stringBuilder.AppendLine()

            For Each item In Body
                stringBuilder.AppendFormat(CultureInfo.CurrentUICulture, "  {0}", New Object(0) {item})
            Next

            stringBuilder.AppendFormat(CultureInfo.CurrentUICulture, "{0}" & VisualBasic.Constants.vbCrLf, New Object(0) {EndLoopToken.Text})
            Return stringBuilder.ToString()
        End Function
    End Class
End Namespace
