Imports System.Globalization
Imports System.Reflection.Emit
Imports Microsoft.SmallVisualBasic.Completion
Imports Microsoft.SmallVisualBasic.Expressions

Namespace Microsoft.SmallVisualBasic.Statements
    Public Class SubroutineCallStatement
        Inherits Statement

        Public Name As Token
        Public Args As List(Of Expression)
        Public IsFunctionCall As Boolean
        Public OuterSubroutine As SubroutineStatement

        Public Overrides Function GetStatementAt(lineNumber As Integer) As Statement
            If lineNumber < StartToken.Line Then Return Nothing
            If Args.Count > 0 AndAlso lineNumber <= Args.Last.EndToken.Line Then Return Me
            Return Nothing
        End Function

        Public Overrides Sub AddSymbols(symbolTable As SymbolTable)
            MyBase.AddSymbols(symbolTable)
            Name.Parent = Me

            For Each arg In Args
                arg.Parent = Me
                arg.AddSymbols(symbolTable)
            Next

            Name.SymbolType = CompletionItemType.SubroutineName
            symbolTable.AddIdentifier(Name)
        End Sub

        Public Overrides Sub EmitIL(scope As CodeGenScope)
            If Args.Count > 0 Then
                Dim code As New Text.StringBuilder()
                For i = Args.Count - 1 To 0 Step -1
                    code.AppendLine($"Stack.PushValue(""_sVB_Arguments"", {Args(i)})")
                Next

                CodeGenerator.LowerAndEmit(code.ToString(), scope, OuterSubroutine, StartToken.Line)
            End If

            Dim methodInfo = scope.MethodBuilders(Name.LCaseText)
            scope.ILGenerator.EmitCall(OpCodes.Call, methodInfo, Nothing)

            If IsFunctionCall Then
                MethodCallStatement.DoNotPopReturnValue = True
                CodeGenerator.LowerAndEmit(
                        $"Stack.PopValue(""_sVB_ReturnValues"")",
                        scope, OuterSubroutine, StartToken.Line)
                MethodCallStatement.DoNotPopReturnValue = False
            End If
        End Sub

        Public Overrides Sub PopulateCompletionItems(
                        bag As CompletionBag,
                        line As Integer,
                        column As Integer,
                        globalScope As Boolean
                   )

            Dim subName = Name.Text
            Dim subName2 = subName.ToLower
            If Name.Contains(line, column, True) Then
                bag.CompletionItems.Add(
                    New CompletionItem() With {
                       .DisplayName = subName,
                       .ItemType = CompletionItemType.SubroutineName,
                       .Key = subName,
                       .ReplacementText = subName,
                       .DefinitionIdintifier = (From subroutine In bag.SymbolTable.Subroutines
                                                Where subroutine.Key = subName2).FirstOrDefault.Value
                    }
                )

            ElseIf Args IsNot Nothing AndAlso Args.Count > 0 AndAlso line >= Args(0).StartToken.Line Then
                CompletionHelper.FillExpressionItems(bag)
                CompletionHelper.FillSubroutines(bag, True)
            End If
        End Sub


        Public Overrides Function ToString() As String
            Dim sb As New Text.StringBuilder(Name.Text)
            Dim n = Args.Count - 1

            sb.Append("(")
            For i = 0 To n
                sb.Append(Args(0).ToString())
                If i < n Then sb.Append(", ")
            Next

            sb.Append(")")
            Return sb.ToString()
        End Function
    End Class
End Namespace
