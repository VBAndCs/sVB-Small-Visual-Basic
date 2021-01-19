Imports System.Collections.Generic
Imports System.Globalization
Imports System.Reflection
Imports System.Reflection.Emit
Imports System.Text
Imports Microsoft.SmallBasic.Completion

Namespace Microsoft.SmallBasic.Statements
    Public Class SubroutineStatement
        Inherits Statement

        Public SubroutineBody As List(Of Statement) = New List(Of Statement)()
        Public SubroutineName As TokenInfo
        Public SubToken As TokenInfo
        Public EndSubToken As TokenInfo

        Public Overrides Sub AddSymbols(ByVal symbolTable As SymbolTable)
            If SubroutineName.Token <> 0 Then
                symbolTable.AddSubroutine(SubroutineName)
            End If

            For Each item In SubroutineBody
                item.AddSymbols(symbolTable)
            Next
        End Sub

        Public Overrides Sub PrepareForEmit(ByVal scope As CodeGenScope)
            Dim methodBuilder = scope.TypeBuilder.DefineMethod(SubroutineName.NormalizedText, MethodAttributes.Static)
            scope.MethodBuilders.Add(SubroutineName.NormalizedText, methodBuilder)
            Dim codeGenScope As CodeGenScope = New CodeGenScope()
            codeGenScope.ILGenerator = methodBuilder.GetILGenerator()
            codeGenScope.TypeBuilder = scope.TypeBuilder
            codeGenScope.MethodBuilder = methodBuilder
            codeGenScope.Parent = scope
            Dim scope2 = codeGenScope

            For Each item In SubroutineBody
                item.PrepareForEmit(scope2)
            Next
        End Sub

        Public Overrides Sub EmitIL(ByVal scope As CodeGenScope)
            Dim methodBuilder = scope.MethodBuilders(SubroutineName.NormalizedText)
            Dim codeGenScope As CodeGenScope = New CodeGenScope()
            codeGenScope.ILGenerator = methodBuilder.GetILGenerator()
            codeGenScope.TypeBuilder = scope.TypeBuilder
            codeGenScope.MethodBuilder = methodBuilder
            codeGenScope.Parent = scope
            Dim codeGenScope2 = codeGenScope

            For Each item In SubroutineBody
                item.EmitIL(codeGenScope2)
            Next

            codeGenScope2.ILGenerator.Emit(OpCodes.Ret)
        End Sub

        Public Overrides Sub PopulateCompletionItems(ByVal completionBag As CompletionBag, ByVal line As Integer, ByVal column As Integer, ByVal globalScope As Boolean)
            If StartToken.Line = line Then
                If SubroutineName.Token = Token.Illegal OrElse column < SubroutineName.Column Then
                    CompletionHelper.FillAllGlobalItems(completionBag, globalScope)
                End If

                Return
            End If

            Dim statementContaining = GetStatementContaining(SubroutineBody, line)

            If statementContaining IsNot Nothing Then
                CompletionHelper.FillKeywords(completionBag, Token.EndSub)
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

            Dim statementContaining = GetStatementContaining(SubroutineBody, line)

            If statementContaining IsNot Nothing Then
                Return 1 + statementContaining.GetIndentationLevel(line)
            End If

            Return 1
        End Function

        Public Overrides Function ToString() As String
            Dim stringBuilder As StringBuilder = New StringBuilder()
            stringBuilder.AppendFormat(CultureInfo.CurrentUICulture, "{0} {1}" & VisualBasic.Constants.vbCrLf, New Object(1) {SubToken.Text, SubroutineName.Text})

            For Each item In SubroutineBody
                stringBuilder.AppendFormat(CultureInfo.CurrentUICulture, "  {0}", New Object(0) {item})
            Next

            stringBuilder.AppendFormat(CultureInfo.CurrentUICulture, "{0}" & VisualBasic.Constants.vbCrLf, New Object(0) {EndSubToken.Text})
            Return stringBuilder.ToString()
        End Function
    End Class
End Namespace
