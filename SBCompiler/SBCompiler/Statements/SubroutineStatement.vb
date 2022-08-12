Imports System.Collections.Generic
Imports System.Globalization
Imports System.Reflection
Imports System.Reflection.Emit
Imports System.Text
Imports Microsoft.SmallBasic.Completion

Namespace Microsoft.SmallBasic.Statements
    Public Class SubroutineStatement
        Inherits Statement

        Friend Shared Current As SubroutineStatement

        Public Name As TokenInfo
        Public Params As List(Of TokenInfo)
        Public Body As New List(Of Statement)()
        Public SubToken As TokenInfo
        Public EndSubToken As TokenInfo
        Friend ReturnStatements As New List(Of ReturnStatement)

        Public Overrides Function GetStatement(lineNumber) As Statement
            If lineNumber < StartToken.Line Then Return Nothing
            If lineNumber > EndSubToken.Line Then Return Nothing

            For Each statment In Body
                Dim st = statment.GetStatement(lineNumber)
                If st IsNot Nothing Then Return st
            Next

            Return Nothing
        End Function

        Public Overrides Function GetKeywords() As LegalTokens
            Dim spans As New LegalTokens
            spans.Add(SubToken)
            For Each statement In ReturnStatements
                spans.Add(statement.StartToken)
            Next
            spans.Add(EndSubToken)
            Return spans
        End Function

        Shared Function GetSubroutine(expression As Expressions.Expression) As SubroutineStatement
            Return GetSubroutine(expression.Parent)
        End Function

        Shared Function GetSubroutine(Token As TokenInfo) As SubroutineStatement
            Return GetSubroutine(Token.Parent)
        End Function

        Shared Function GetSubroutine(statement As Statement) As SubroutineStatement
            Do Until statement Is Nothing OrElse TypeOf statement Is SubroutineStatement
                statement = statement.Parent
            Loop
            Return statement
        End Function

        Public Overrides Sub AddSymbols(ByVal symbolTable As SymbolTable)
            MyBase.AddSymbols(symbolTable)
            Name.Parent = Me
            SubToken.Parent = Me
            EndSubToken.Parent = Me

            If Params IsNot Nothing Then
                For Each param In Params
                    param.Parent = Me
                Next
            End If
            If Name.Token <> 0 Then
                symbolTable.AddSubroutine(Name, StartToken.Token)
            End If

            For Each item In Body
                item.Parent = Me
                item.AddSymbols(symbolTable)
            Next
        End Sub

        Public Overrides Sub PrepareForEmit(scope As CodeGenScope)
            Dim methodBuilder = scope.TypeBuilder.DefineMethod(Name.NormalizedText, MethodAttributes.Static)
            scope.MethodBuilders.Add(Name.NormalizedText, methodBuilder)
            Dim codeGenScope As New CodeGenScope() With {
                .ILGenerator = methodBuilder.GetILGenerator(),
                .TypeBuilder = scope.TypeBuilder,
                .MethodBuilder = methodBuilder,
                .Parent = scope
            }

            For Each item In Body
                item.PrepareForEmit(codeGenScope)
            Next
        End Sub

        Public Overrides Sub EmitIL(ByVal scope As CodeGenScope)
            Dim methodBuilder = scope.MethodBuilders(Name.NormalizedText)
            Dim codeGenScope As New CodeGenScope() With {
                .ILGenerator = methodBuilder.GetILGenerator(),
                .TypeBuilder = scope.TypeBuilder,
                .MethodBuilder = methodBuilder,
                .Parent = scope
            }

            For Each item In Body
                item.EmitIL(codeGenScope)
            Next

            codeGenScope.ILGenerator.Emit(OpCodes.Ret)
        End Sub

        Public Overrides Sub PopulateCompletionItems(
                         completionBag As CompletionBag,
                         line As Integer,
                         column As Integer,
                         globalScope As Boolean)

            If StartToken.Line = line Then
                If Name.Token = Token.Illegal OrElse column < Name.Column Then
                    CompletionHelper.FillAllGlobalItems(completionBag, globalScope)
                End If
                Return
            End If

            Dim statementContaining = GetStatementContaining(Body, line)
            CompletionHelper.FillLocals(completionBag, Name.NormalizedText)

            If statementContaining IsNot Nothing Then
                If StartToken.Token = Token.Sub Then
                    CompletionHelper.FillKeywords(completionBag, Token.EndSub)
                Else
                    CompletionHelper.FillKeywords(completionBag, Token.EndFunction)
                End If

                statementContaining.PopulateCompletionItems(completionBag, line, column, globalScope:=False)
            End If
        End Sub

        Public Overrides Function ToString() As String
            Dim stringBuilder As StringBuilder = New StringBuilder()
            stringBuilder.AppendFormat(CultureInfo.CurrentUICulture, "{0} {1}" & VisualBasic.Constants.vbCrLf, New Object(1) {SubToken.Text, Name.Text})

            For Each item In Body
                stringBuilder.AppendFormat(CultureInfo.CurrentUICulture, "  {0}", New Object(0) {item})
            Next

            stringBuilder.AppendFormat(CultureInfo.CurrentUICulture, "{0}" & VisualBasic.Constants.vbCrLf, New Object(0) {EndSubToken.Text})
            Return stringBuilder.ToString()
        End Function

    End Class
End Namespace
