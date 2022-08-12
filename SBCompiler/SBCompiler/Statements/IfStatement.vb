Imports System.Collections.Generic
Imports System.Globalization
Imports System.Reflection.Emit
Imports System.Text
Imports Microsoft.SmallBasic.Completion
Imports Microsoft.SmallBasic.Expressions

Namespace Microsoft.SmallBasic.Statements
    Public Class IfStatement
        Inherits Statement

        Public ThenStatements As New List(Of Statement)()
        Public ElseIfStatements As New List(Of ElseIfStatement)()
        Public ElseStatements As New List(Of Statement)()
        Public Condition As Expression
        Public IfToken As TokenInfo
        Public ThenToken As TokenInfo
        Public ElseToken As TokenInfo
        Public EndIfToken As TokenInfo

        Public Overrides Function GetStatement(lineNumber) As Statement
            If lineNumber < StartToken.Line Then Return Nothing
            If lineNumber <= ThenToken.Line Then Return Me
            If lineNumber > EndIfToken.Line Then Return Nothing

            For Each statment In ThenStatements
                Dim st = statment.GetStatement(lineNumber)
                If st IsNot Nothing Then Return st
            Next

            For Each statment In ElseIfStatements
                Dim st = statment.GetStatement(lineNumber)
                If st IsNot Nothing Then Return st
            Next

            For Each statment In ElseStatements
                Dim st = statment.GetStatement(lineNumber)
                If st IsNot Nothing Then Return st
            Next

            If lineNumber <= EndIfToken.Line Then Return Me

            Return Nothing
        End Function



        Public Overrides Function GetKeywords() As LegalTokens
            Dim spans As New LegalTokens
            spans.Add(IfToken)
            spans.Add(ThenToken)

            For Each statement In ElseIfStatements
                spans.Add(statement.ElseIfToken)
                spans.Add(statement.ThenToken)
            Next

            spans.Add(ElseToken)
            spans.Add(EndIfToken)
            Return spans
        End Function


        Public Overrides Sub AddSymbols(symbolTable As SymbolTable)
            MyBase.AddSymbols(symbolTable)
            IfToken.Parent = Me
            ThenToken.Parent = Me
            ElseToken.Parent = Me
            EndIfToken.Parent = Me

            If Condition IsNot Nothing Then
                Condition.Parent = Me
                Condition.AddSymbols(symbolTable)
            End If

            For Each thenStatement In ThenStatements
                thenStatement.Parent = Me
                thenStatement.AddSymbols(symbolTable)
            Next

            For Each elseIfStatement In ElseIfStatements
                elseIfStatement.Parent = Me
                elseIfStatement.AddSymbols(symbolTable)
            Next

            For Each elseStatement In ElseStatements
                elseStatement.Parent = Me
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
            Dim endIfLabel As Label = scope.ILGenerator.DefineLabel()
            Dim elseIfLabel As Label = scope.ILGenerator.DefineLabel()
            Condition.EmitIL(scope)
            scope.ILGenerator.EmitCall(OpCodes.Call, scope.TypeInfoBag.PrimitiveToBoolean, Nothing)
            scope.ILGenerator.Emit(OpCodes.Brfalse, elseIfLabel)

            For Each thenStatement In ThenStatements
                thenStatement.EmitIL(scope)
            Next

            scope.ILGenerator.Emit(OpCodes.Br, endIfLabel)

            For Each elseIfStatement In ElseIfStatements
                scope.ILGenerator.MarkLabel(elseIfLabel)
                elseIfLabel = scope.ILGenerator.DefineLabel()
                elseIfStatement.Condition.EmitIL(scope)
                scope.ILGenerator.EmitCall(OpCodes.Call, scope.TypeInfoBag.PrimitiveToBoolean, Nothing)
                scope.ILGenerator.Emit(OpCodes.Brfalse, elseIfLabel)

                For Each thenStatement2 In elseIfStatement.ThenStatements
                    thenStatement2.EmitIL(scope)
                Next

                scope.ILGenerator.Emit(OpCodes.Br, endIfLabel)
            Next

            scope.ILGenerator.MarkLabel(elseIfLabel)

            For Each elseStatement In ElseStatements
                elseStatement.EmitIL(scope)
            Next

            scope.ILGenerator.MarkLabel(endIfLabel)
        End Sub

        Public Overrides Sub PopulateCompletionItems(
                                         completionBag As CompletionBag,
                                         line As Integer,
                                         column As Integer,
                                         globalScope As Boolean
                                  )

            Select Case line
                Case StartToken.Line
                    If column < IfToken.EndColumn Then
                        CompletionHelper.FillAllGlobalItems(completionBag, globalScope)
                    Else
                        CompletionHelper.FillLogicalExpressionItems(completionBag)
                        CompletionHelper.FillKeywords(completionBag, Token.Then)
                    End If

                Case ElseToken.Line
                    CompletionHelper.FillKeywords(completionBag, Token.Else, Token.EndIf)
                    CompletionHelper.FillAllGlobalItems(completionBag, inGlobalScope:=False)

                Case EndIfToken.Line
                    CompletionHelper.FillKeywords(completionBag, Token.EndIf)
                    CompletionHelper.FillAllGlobalItems(completionBag, inGlobalScope:=False)

                Case Else
                    CompletionHelper.FillKeywords(completionBag, Token.EndIf, Token.Else, Token.ElseIf)
                    Dim statementContaining = GetStatementContaining(ElseStatements, line)
                    If statementContaining Is Nothing Then
                        statementContaining = GetStatementContaining(ElseIfStatements.OfType(Of Statement).ToList(), line)
                        If statementContaining Is Nothing Then
                            statementContaining = GetStatementContaining(ThenStatements, line)
                        End If
                    End If
                    statementContaining?.PopulateCompletionItems(completionBag, line, column, globalScope:=False)
            End Select
        End Sub


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
