Imports System.Collections.Generic
Imports System.Globalization
Imports System.Reflection.Emit
Imports System.Text
Imports Microsoft.SmallBasic.Completion
Imports Microsoft.SmallBasic.Expressions

Namespace Microsoft.SmallBasic.Statements
    Public Class IfStatement
        Inherits Statement

        Public ThenStatements As List(Of Statement) = New List(Of Statement)()
        Public ElseIfStatements As List(Of Statement) = New List(Of Statement)()
        Public ElseStatements As List(Of Statement) = New List(Of Statement)()
        Public Condition As Expression
        Public IfToken As TokenInfo
        Public ThenToken As TokenInfo
        Public ElseToken As TokenInfo
        Public EndIfToken As TokenInfo

        Public Overrides Sub AddSymbols(ByVal symbolTable As SymbolTable)
            If Condition IsNot Nothing Then
                Condition.AddSymbols(symbolTable)
            End If

            For Each thenStatement In ThenStatements
                thenStatement.AddSymbols(symbolTable)
            Next

            For Each elseIfStatement In ElseIfStatements
                elseIfStatement.AddSymbols(symbolTable)
            Next

            For Each elseStatement In ElseStatements
                elseStatement.AddSymbols(symbolTable)
            Next
        End Sub

        Public Overrides Sub PrepareForEmit(ByVal scope As CodeGenScope)
            For Each thenStatement In ThenStatements
                thenStatement.PrepareForEmit(scope)
            Next

            For Each elseIfStatement In ElseIfStatements
                elseIfStatement.PrepareForEmit(scope)
            Next

            For Each elseStatement In ElseStatements
                elseStatement.PrepareForEmit(scope)
            Next
        End Sub

        Public Overrides Sub EmitIL(ByVal scope As CodeGenScope)
            Dim label As Label = scope.ILGenerator.DefineLabel()
            Dim label2 As Label = scope.ILGenerator.DefineLabel()
            Condition.EmitIL(scope)
            scope.ILGenerator.EmitCall(OpCodes.Call, scope.TypeInfoBag.PrimitiveToBoolean, Nothing)
            scope.ILGenerator.Emit(OpCodes.Brfalse, label2)

            For Each thenStatement In ThenStatements
                thenStatement.EmitIL(scope)
            Next

            scope.ILGenerator.Emit(OpCodes.Br, label)

            For Each elseIfStatement As ElseIfStatement In ElseIfStatements
                scope.ILGenerator.MarkLabel(label2)
                label2 = scope.ILGenerator.DefineLabel()
                elseIfStatement.Condition.EmitIL(scope)
                scope.ILGenerator.EmitCall(OpCodes.Call, scope.TypeInfoBag.PrimitiveToBoolean, Nothing)
                scope.ILGenerator.Emit(OpCodes.Brfalse, label2)

                For Each thenStatement2 In elseIfStatement.ThenStatements
                    thenStatement2.EmitIL(scope)
                Next

                scope.ILGenerator.Emit(OpCodes.Br, label)
            Next

            scope.ILGenerator.MarkLabel(label2)

            For Each elseStatement In ElseStatements
                elseStatement.EmitIL(scope)
            Next

            scope.ILGenerator.MarkLabel(label)
        End Sub

        Public Overrides Sub PopulateCompletionItems(ByVal completionBag As CompletionBag, ByVal line As Integer, ByVal column As Integer, ByVal globalScope As Boolean)
            If StartToken.Line = line Then
                If column < IfToken.EndColumn Then
                    CompletionHelper.FillAllGlobalItems(completionBag, globalScope)
                    Return
                End If

                CompletionHelper.FillLogicalExpressionItems(completionBag)
                CompletionHelper.FillKeywords(completionBag, Token.Then)
                Return
            End If

            If line = ElseToken.Line Then
                CompletionHelper.FillKeywords(completionBag, Token.Else, Token.EndIf)
                CompletionHelper.FillAllGlobalItems(completionBag, inGlobalScope:=False)
                Return
            End If

            If line = EndIfToken.Line Then
                CompletionHelper.FillKeywords(completionBag, Token.EndIf)
                CompletionHelper.FillAllGlobalItems(completionBag, inGlobalScope:=False)
                Return
            End If

            Dim statementContaining = GetStatementContaining(ElseStatements, line)
            CompletionHelper.FillKeywords(completionBag, Token.EndIf)

            If statementContaining Is Nothing Then
                CompletionHelper.FillKeywords(completionBag, Token.Else)
                statementContaining = GetStatementContaining(ElseIfStatements, line)

                If statementContaining Is Nothing Then
                    CompletionHelper.FillKeywords(completionBag, Token.ElseIf)
                    statementContaining = GetStatementContaining(ThenStatements, line)
                End If
            End If

            statementContaining?.PopulateCompletionItems(completionBag, line, column, globalScope:=False)
        End Sub

        Public Overrides Function GetIndentationLevel(ByVal line As Integer) As Integer
            If line = IfToken.Line Then
                Return 0
            End If

            If ElseToken.Token <> 0 AndAlso line = ElseToken.Line Then
                Return 0
            End If

            If EndIfToken.Token <> 0 AndAlso line = EndIfToken.Line Then
                Return 0
            End If

            Dim statementContaining = GetStatementContaining(ElseStatements, line)

            If statementContaining IsNot Nothing Then
                Return 1 + statementContaining.GetIndentationLevel(line)
            End If

            statementContaining = GetStatementContaining(ElseIfStatements, line)

            If statementContaining IsNot Nothing Then
                Return statementContaining.GetIndentationLevel(line)
            End If

            statementContaining = GetStatementContaining(ThenStatements, line)

            If statementContaining IsNot Nothing Then
                Return 1 + statementContaining.GetIndentationLevel(line)
            End If

            Return 1
        End Function

        Public Overrides Function ToString() As String
            Dim stringBuilder As StringBuilder = New StringBuilder()
            stringBuilder.AppendFormat(CultureInfo.CurrentUICulture, "{0} {1} {2}" & VisualBasic.Constants.vbCrLf, New Object(2) {IfToken.Text, Condition, ThenToken.Text})

            For Each thenStatement In ThenStatements
                stringBuilder.AppendFormat(CultureInfo.CurrentUICulture, "  {0}", New Object(0) {thenStatement})
            Next

            For Each elseIfStatement In ElseIfStatements
                stringBuilder.AppendFormat(CultureInfo.CurrentUICulture, "  {0}", New Object(0) {elseIfStatement})
            Next

            If ElseToken.TokenType <> 0 Then
                stringBuilder.AppendFormat(CultureInfo.CurrentUICulture, "{0}" & VisualBasic.Constants.vbCrLf, New Object(0) {ElseToken.Text})

                For Each elseStatement In ElseStatements
                    stringBuilder.AppendFormat(CultureInfo.CurrentUICulture, "  {0}", New Object(0) {elseStatement})
                Next
            End If

            stringBuilder.AppendFormat(CultureInfo.CurrentUICulture, "{0}" & VisualBasic.Constants.vbCrLf, New Object(0) {EndIfToken.Text})
            Return stringBuilder.ToString()
        End Function
    End Class
End Namespace
