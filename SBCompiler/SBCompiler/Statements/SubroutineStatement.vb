Imports System.Reflection
Imports System.Reflection.Emit
Imports System.Text
Imports Microsoft.SmallVisualBasic.Completion

Namespace Microsoft.SmallVisualBasic.Statements
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
            If lineNumber = SubToken.Line Then Return Me
            If lineNumber = EndSubToken.Line Then Return Me
            If lineNumber < StartToken.Line Then Return Nothing
            If lineNumber > EndSubToken.Line Then Return Nothing

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
                FillParamsComments()
                For Each param In Params
                    param.Parent = Me
                    param.SymbolType = CompletionItemType.LocalVariable
                    Dim paramId As New Expressions.IdentifierExpression() With {
                        .Identifier = param,
                        .Subroutine = Me
                    }
                    Dim key = symbolTable.AddVariable(paramId, param.Comment, True)
                    symbolTable.InferedTypes(key) = WinForms.PreCompiler.GetVarType(param.Text)
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
                    id.SymbolType = CompletionItemType.LocalVariable
                    ids(i) = id
                End If
            Next

        End Sub

        Public Overrides Sub PrepareForEmit(scope As CodeGenScope)
            If scope.ForGlobalHelp AndAlso Name.IsIllegal Then Return

            Dim methodBuilder = scope.TypeBuilder.DefineMethod(
                Name.Text,
                MethodAttributes.Static Or MethodAttributes.Public
            )

            Dim prmtvType = GetType(Library.Primitive)
            Dim n = If(Params IsNot Nothing, Params.Count - 1, -1)

            If n > -1 Then
                Dim paramTypes(n) As Type
                For i = 0 To n
                    paramTypes(i) = prmtvType
                Next
                methodBuilder.SetParameters(paramTypes)
            End If

            If SubToken.Type = TokenType.Function Then
                methodBuilder.SetReturnType(prmtvType)
            Else
                methodBuilder.SetReturnType(GetType(Void))
            End If

            methodBuilder.DefineParameter(0, ParameterAttributes.None, "")

            For i = 0 To n
                methodBuilder.DefineParameter(
                     i + 1,
                     ParameterAttributes.None,
                     Params(i).Text
                )
            Next

            Dim returntype = InferReturnType(scope.SymbolTable)
            If returntype <> VariableType.Any Then
                Dim ctorParams = New Type() {GetType(VariableType)}
                Dim ctorInfo = GetType(WinForms.ReturnValueTypeAttribute).GetConstructor(ctorParams)
                methodBuilder.SetCustomAttribute(
                    New CustomAttributeBuilder(
                            ctorInfo,
                            New Object() {returntype}
                    )
                )
            End If

            scope.MethodBuilders.Add(Name.LCaseText, methodBuilder)

            Dim codeGenScope As New CodeGenScope() With {
                .ILGenerator = methodBuilder.GetILGenerator(),
                .TypeBuilder = scope.TypeBuilder,
                .MethodBuilder = methodBuilder,
                .Parent = scope
            }

            If scope.ForGlobalHelp Then Return

            For Each item In Body
                item.PrepareForEmit(codeGenScope)
            Next
        End Sub

        Public Overrides Sub EmitIL(scope As CodeGenScope)
            If scope.ForGlobalHelp AndAlso Name.IsIllegal Then Return

            Dim methodBuilder = scope.MethodBuilders(Name.LCaseText)
            Dim codeGenScope As New CodeGenScope() With {
                .ILGenerator = methodBuilder.GetILGenerator(),
                .TypeBuilder = scope.TypeBuilder,
                .MethodBuilder = methodBuilder,
                .Parent = scope
            }

            If Not scope.ForGlobalHelp Then
                If Params IsNot Nothing Then
                    For i = 0 To Params.Count - 1
                        codeGenScope.ILGenerator.Emit(OpCodes.Ldarg, i)
                        Dim var = codeGenScope.GetLocalBuilder(Me, Params(i))
                        codeGenScope.ILGenerator.Emit(OpCodes.Stloc, var)
                    Next
                End If

                For Each item In Body
                    item.EmitIL(codeGenScope)
                Next
            End If

            If SubToken.Type = TokenType.Function Then
                ' return zero just in case not all pathes of code has a return value
                Expressions.LiteralExpression.Zero.EmitIL(codeGenScope)
            End If
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

            ElseIf Params IsNot Nothing AndAlso Params.Count > 0 AndAlso (
                           line >= Params(0).Line AndAlso
                           line <= Params.Last.Line
                      ) Then

                If bag.ForHelp Then
                    For Each param In Params
                        bag.CompletionItems.Add(
                            New Completion.CompletionItem() With {
                                .DisplayName = param.Text,
                                .Key = Name.LCaseText & "." & param.LCaseText,
                                .ItemType = Completion.CompletionItemType.LocalVariable,
                                .DefinitionIdintifier = param
                            })
                    Next
                Else
                    bag.ShowCompletion = False
                End If
            Else
                CompletionHelper.FillLocals(bag, Name.LCaseText)
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
                    sb.Append(Params(i).Text)
                    If i < n Then sb.Append(", ")
                Next
                sb.AppendLine(")")
            End If

            For Each st In Body
                sb.Append(st.ToString())
            Next

            sb.AppendLine(EndSubToken.Text)
            Return sb.ToString()
        End Function

        Public Overloads Function InferReturnType(symbolTable As SymbolTable) As VariableType
            Dim containsString = False
            Dim containsControl = False
            Dim sameType = True
            Dim returnType = VariableType.Any

            For Each retSt In ReturnStatements
                Dim type = retSt.ReturnExpression.InferType(symbolTable)
                If type <> VariableType.Any Then
                    Select Case type
                        Case VariableType.String, VariableType.Array,
                             VariableType.Color, >= VariableType.Control
                            containsString = True
                        Case >= VariableType.Control
                            containsControl = True
                    End Select

                    If returnType <> VariableType.Any Then
                        If type <> returnType Then sameType = False
                    End If
                    returnType = type
                End If
            Next

            If sameType Then Return returnType
            If containsString Then Return VariableType.String
            If containsControl Then Return VariableType.Control
            Return VariableType.Double
        End Function

        Public Overrides Sub InferType(symbolTable As SymbolTable)
            For Each st In Body
                st.InferType(symbolTable)
            Next
        End Sub
    End Class
End Namespace
