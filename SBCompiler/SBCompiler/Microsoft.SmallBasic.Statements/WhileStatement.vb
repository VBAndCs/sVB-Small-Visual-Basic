Imports System.Collections.Generic
Imports System.Globalization
Imports System.Reflection.Emit
Imports System.Text
Imports Microsoft.SmallBasic.Completion
Imports Microsoft.SmallBasic.Expressions

Namespace Microsoft.SmallBasic.Statements
    Public Class WhileStatement
        Inherits LoopStatement

        Public WhileBody As List(Of Statement) = New List(Of Statement)()
        Public Condition As Expression
        Public WhileToken As TokenInfo
        Public EndWhileToken As TokenInfo

        Public Overrides Sub AddSymbols(ByVal symbolTable As SymbolTable)
            If Condition IsNot Nothing Then
                Condition.Parent = Me
                Condition.AddSymbols(symbolTable)
            End If

            For Each item In WhileBody
                item.Parent = Me
                item.AddSymbols(symbolTable)
            Next
        End Sub

        Public Overrides Sub PrepareForEmit(ByVal scope As CodeGenScope)
            For Each item In WhileBody
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

            For Each item In WhileBody
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

            Dim statementContaining = GetStatementContaining(WhileBody, line)

            If statementContaining IsNot Nothing Then
                CompletionHelper.FillKeywords(completionBag, Token.EndWhile, Token.Wend)
                statementContaining.PopulateCompletionItems(completionBag, line, column, globalScope:=False)
            End If
        End Sub

        Public Overrides Function GetIndentationLevel(ByVal line As Integer) As Integer
            If line = WhileToken.Line Then
                Return 0
            End If

            If EndWhileToken.Token <> 0 AndAlso line = EndWhileToken.Line Then
                Return 0
            End If

            Dim statementContaining = GetStatementContaining(WhileBody, line)

            If statementContaining IsNot Nothing Then
                Return 1 + statementContaining.GetIndentationLevel(line)
            End If

            Return 1
        End Function

        Public Overrides Function ToString() As String
            Dim stringBuilder As StringBuilder = New StringBuilder()
            stringBuilder.AppendFormat(CultureInfo.CurrentUICulture, "{0} {1}", New Object(1) {WhileToken.Text, Condition})
            stringBuilder.AppendLine()

            For Each item In WhileBody
                stringBuilder.AppendFormat(CultureInfo.CurrentUICulture, "  {0}", New Object(0) {item})
            Next

            stringBuilder.AppendFormat(CultureInfo.CurrentUICulture, "{0}" & VisualBasic.Constants.vbCrLf, New Object(0) {EndWhileToken.Text})
            Return stringBuilder.ToString()
        End Function
    End Class
End Namespace
