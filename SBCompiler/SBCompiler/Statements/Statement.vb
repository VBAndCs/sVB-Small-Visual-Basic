Imports System.Collections.Generic
Imports Microsoft.SmallVisualBasic.Completion

Namespace Microsoft.SmallVisualBasic.Statements
    Public MustInherit Class Statement
        Public StartToken As Token

        Public EndingComment As Token
        Public LeadingComment As String
        Public Parent As Statement

        Dim sepLineChars() As Char = {"-"c, "_"c, "*", "#", "$"c, "&"c, "+"c, "="c, "!"c}

        Public Overridable Sub AddSymbols(symbolTable As SymbolTable)
            StartToken.Parent = Me
            EndingComment.Parent = Me
            If TypeOf Me Is EmptyStatement Then Return

            If Not symbolTable.IsLoweredCode AndAlso Not symbolTable.AllStatements.ContainsKey(StartToken.Line) Then
                symbolTable.AllStatements(StartToken.Line) = Me
            End If

            Dim comments = symbolTable.AllCommentLines
            Dim line = StartToken.Line - 1
            Dim foundComments As New List(Of String)

            For i = comments.Count - 1 To 0 Step -1
                Dim commentLine = comments(i).Line
                If commentLine = line Then
                    Dim text = comments(i).Text.TrimStart("'"c, " "c, vbTab)
                    If text.Trim(sepLineChars) = "" Then Exit For
                    line -= 1
                    foundComments.Add(text)

                ElseIf commentLine < line Then
                    Exit For
                End If
            Next

            Dim sb As New Text.StringBuilder
            For i = foundComments.Count - 1 To 0 Step -1
                sb.Append(foundComments(i))
                If i > 0 Then sb.AppendLine()
            Next
            LeadingComment = sb.ToString()
        End Sub

        Public Overridable Function GetSummery() As String
            If EndingComment.Type = TokenType.Illegal Then Return LeadingComment
            Dim endCom = EndingComment.Text.Trim("'"c, " "c, vbTab)
            If LeadingComment = "" Then Return endCom
            Return LeadingComment & vbCrLf & endCom
        End Function

        Public MustOverride Function GetStatementAt(lineNumber As Integer) As Statement

        Public Overridable Sub PrepareForEmit(scope As CodeGenScope)
        End Sub

        Public Overridable Sub EmitIL(scope As CodeGenScope)
        End Sub

        Public Overridable Sub PopulateCompletionItems(completionBag As CompletionBag, line As Integer, column As Integer, globalScope As Boolean)
            CompletionHelper.FillAllGlobalItems(completionBag, globalScope)
        End Sub

        Public Shared Function GetStatementContaining(statements As List(Of Statement), line As Integer) As Statement
            For num = statements.Count - 1 To 0 Step -1
                Dim statement = statements(num)
                Dim L = statement.StartToken.Line
                If L > 0 AndAlso line >= L Then Return statement
            Next
            Return Nothing
        End Function

        Public Overridable Function GetKeywords() As LegalTokens
            Return Nothing
        End Function

        Public Function IsOfType(statementType As Type) As Boolean
            If statementType Is Nothing Then
                Select Case Me.GetType
                    Case GetType(SubroutineStatement),
                             GetType(IfStatement),
                             GetType(ForStatement),
                             GetType(WhileStatement)
                        Return True
                    Case Else
                        Return False
                End Select
            End If
            Return Me.GetType Is statementType
        End Function

        Public Overridable Sub InferType(symbolTable As SymbolTable)
            Select Case Me.GetType()
                Case GetType(AssignmentStatement)
                    CType(Me, AssignmentStatement).InferType(symbolTable)

                Case GetType(IfStatement)
                    CType(Me, IfStatement).InferType(symbolTable)

                Case GetType(ForEachStatement)
                    CType(Me, ForEachStatement).InferType(symbolTable)

                Case GetType(ForStatement)
                    CType(Me, ForStatement).InferType(symbolTable)

                Case GetType(WhileStatement)
                    CType(Me, WhileStatement).InferType(symbolTable)

                Case GetType(SubroutineStatement)
                    CType(Me, SubroutineStatement).InferType(symbolTable)

            End Select
        End Sub
    End Class
End Namespace
