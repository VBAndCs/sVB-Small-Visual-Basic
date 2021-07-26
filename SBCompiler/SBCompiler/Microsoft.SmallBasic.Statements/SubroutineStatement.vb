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

        Public Overrides Sub AddSymbols(ByVal symbolTable As SymbolTable)
            If Name.Token <> 0 Then
                symbolTable.AddSubroutine(Name, StartToken.Token)
            End If

            For Each item In Body
                item.Parent = Me
                item.AddSymbols(symbolTable)
            Next
        End Sub

        Public Overrides Sub PrepareForEmit(ByVal scope As CodeGenScope)
            Dim methodBuilder = scope.TypeBuilder.DefineMethod(Name.NormalizedText, MethodAttributes.Static)
            scope.MethodBuilders.Add(Name.NormalizedText, methodBuilder)
            Dim codeGenScope As CodeGenScope = New CodeGenScope()
            codeGenScope.ILGenerator = methodBuilder.GetILGenerator()
            codeGenScope.TypeBuilder = scope.TypeBuilder
            codeGenScope.MethodBuilder = methodBuilder
            codeGenScope.Parent = scope
            Dim scope2 = codeGenScope

            For Each item In Body
                item.PrepareForEmit(scope2)
            Next
        End Sub

        Public Overrides Sub EmitIL(ByVal scope As CodeGenScope)
            Dim methodBuilder = scope.MethodBuilders(Name.NormalizedText)
            Dim codeGenScope As New CodeGenScope()
            codeGenScope.ILGenerator = methodBuilder.GetILGenerator()
            codeGenScope.TypeBuilder = scope.TypeBuilder
            codeGenScope.MethodBuilder = methodBuilder
            codeGenScope.Parent = scope
            Dim codeGenScope2 = codeGenScope

            For Each item In Body
                item.EmitIL(codeGenScope2)
            Next

            codeGenScope2.ILGenerator.Emit(OpCodes.Ret)
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

        Public Overrides Function GetIndentationLevel(ByVal line As Integer) As Integer
            If line = SubToken.Line Then
                Return 0
            End If

            If EndSubToken.Token <> 0 AndAlso line = EndSubToken.Line Then
                Return 0
            End If

            Dim statementContaining = GetStatementContaining(Body, line)

            If statementContaining IsNot Nothing Then
                Return 1 + statementContaining.GetIndentationLevel(line)
            End If

            Return 1
        End Function

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
