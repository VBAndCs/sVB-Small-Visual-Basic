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

        Public SubToken As Token
        Public Name As Token
        Public Params As List(Of Token)
        Public Body As New List(Of Statement)()
        Public EndSubToken As Token
        Friend ReturnStatements As New List(Of ReturnStatement)
        Public SubLineComments As List(Of Token)

        Public Sub New()
            If LineScanner.SubLineComments.Count > 0 Then
                SubLineComments = New List(Of Token)
                SubLineComments.AddRange(LineScanner.SubLineComments)
            End If
        End Sub

        Public Overrides Function GetSummery() As String
            Dim comment As New StringBuilder()
            comment.AppendLine(LeadingComment)

            If SubLineComments IsNot Nothing AndAlso
                    SubLineComments(0).subLine = 0 AndAlso
                     (Params?.Count = 0 OrElse Params(0).subLine > 0) Then
                comment.AppendLine(SubLineComments(0).Text.TrimStart("'"c, " "c, vbTab))
            End If

            If SubToken.Type = TokenType.Sub AndAlso EndingComment.Type <> TokenType.Illegal Then
                comment.AppendLine(EndingComment.Text.TrimStart("'"c, " "c, vbTab))
            End If

            Return comment.ToString().Trim()
        End Function

        Public Function GetRetunDoc() As String
            If SubToken.Type = TokenType.Sub OrElse
                    EndingComment.Type = TokenType.Illegal Then Return ""

            Return EndingComment.Text.TrimStart("'"c, " "c)
        End Function

        Dim fillComments As Boolean = True

        Public Function GetParamsDoc() As Dictionary(Of String, String)
            If Params Is Nothing Then Return Nothing

            Dim docs As New Dictionary(Of String, String)

            For Each param In Params
                docs.Add(param.Text, param.Comment)
            Next

            Return docs
        End Function

        Friend Sub FillParamsComments()
            If fillComments Then
                fillComments = False
            Else
                Return
            End If

            If SubLineComments Is Nothing Then Return

            For i = SubLineComments.Count - 1 To 0 Step -1
                For j = Params.Count - 1 To 0 Step -1
                    Dim subline = SubLineComments(i).subLine
                    Dim p = Params(j)
                    If p.subLine = subline Then
                        p.Comment = SubLineComments(i).Text.TrimStart("'"c, " "c)
                        Params(j) = p
                        Exit For
                    End If
                Next
            Next
        End Sub


        Public Overrides Function GetStatementAt(lineNumber As Integer) As Statement
            If lineNumber < StartToken.Line Then Return Nothing
            If lineNumber > EndSubToken.Line Then Return Nothing
            If lineNumber = EndSubToken.Line Then Return Me

            For Each statment In Body
                Dim st = statment.GetStatementAt(lineNumber)
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

        Shared Function GetSubroutine(Token As Token) As SubroutineStatement
            Return GetSubroutine(Token.Parent)
        End Function

        Shared Function GetSubroutine(statement As Statement) As SubroutineStatement
            Do Until statement Is Nothing OrElse TypeOf statement Is SubroutineStatement
                statement = statement.Parent
            Loop
            Return statement
        End Function

        Public Overrides Sub AddSymbols(symbolTable As SymbolTable)
            MyBase.AddSymbols(symbolTable)

            Name.Parent = Me
            SubToken.Parent = Me
            EndSubToken.Parent = Me

            If Params IsNot Nothing Then
                For Each param In Params
                    param.Parent = Me
                    param.SymbolType = CompletionItemType.LocalVariable
                    symbolTable.AddIdentifier(param)
                Next
            End If

            If Not Name.IsIllegal Then
                symbolTable.AddSubroutine(Name, StartToken.Type)
                Name.SymbolType = CompletionItemType.SubroutineName
                symbolTable.AddIdentifier(Name)
            End If

            For Each item In Body
                item.Parent = Me
                item.AddSymbols(symbolTable)
            Next

            ' Fix local variables type, because For loops can make a var local after it was gloval in preb lines!
            Dim ids = symbolTable.AllIdentifiers
            For i = 0 To ids.Count - 1
                Dim id = ids(i)
                If id.SymbolType = CompletionItemType.GlobalVariable AndAlso symbolTable.IsLocalVar(id) Then
                    id.SymbolType = Completion.CompletionItemType.LocalVariable
                    ids(i) = id
                End If
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

        Public Overrides Sub EmitIL(scope As CodeGenScope)
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
                         bag As CompletionBag,
                         line As Integer,
                         column As Integer,
                         globalScope As Boolean)

            If StartToken.Line = line AndAlso Name.IsAfter(line, column) Then
                CompletionHelper.FillAllGlobalItems(bag, globalScope)

            ElseIf Name.Contains(line, column, bag.ForHelp) Then
                bag.CompletionItems.Add(
                            New Completion.CompletionItem() With {
                                     .DisplayName = Name.Text,
                                     .Key = Name.Text,
                                     .ItemType = Completion.CompletionItemType.SubroutineName,
                                     .DefinitionIdintifier = Name
                    })

            ElseIf Params IsNot Nothing AndAlso Params.count > 0 AndAlso (
                           line >= Params(0).Line AndAlso
                           line <= Params.Last.Line
                      ) Then

                If bag.ForHelp Then
                    For Each param In Params
                        bag.CompletionItems.Add(
                            New Completion.CompletionItem() With {
                                .DisplayName = param.Text,
                                .Key = Name.NormalizedText & "." & param.NormalizedText,
                                .ItemType = Completion.CompletionItemType.LocalVariable,
                                .DefinitionIdintifier = param
                            })
                    Next
                Else
                    bag.ShowCompletion = False
                End If
            Else
                CompletionHelper.FillLocals(bag, Name.NormalizedText)
                Dim statement = GetStatementContaining(Body, line)
                If statement IsNot Nothing Then
                    If StartToken.Type = TokenType.Sub Then
                        CompletionHelper.FillKeywords(bag, {TokenType.EndSub})
                    Else
                        CompletionHelper.FillKeywords(bag, {TokenType.EndFunction})
                    End If

                    statement.PopulateCompletionItems(bag, line, column, globalScope:=False)
                End If
            End If
        End Sub

        Public Overrides Function ToString() As String
            Dim sb As New StringBuilder(SubToken.Text)
            sb.Append(" ")
            sb.Append(Name.Text)
            If Params IsNot Nothing Then
                Dim n = Params.Count - 1
                sb.Append("(")

                For i = 0 To n
                    sb.Append(Params(0).ToString())
                    If i < n Then sb.Append(", ")
                Next

                sb.AppendLine(")")
            End If

            For Each item In Body
                sb.AppendLine(item.ToString())
            Next

            sb.AppendLine(EndSubToken.Text)
            Return sb.ToString()
        End Function


    End Class
End Namespace
