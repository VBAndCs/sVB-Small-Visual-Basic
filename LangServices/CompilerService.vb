Imports System.IO
Imports Microsoft.Nautilus.Text
Imports Microsoft.SmallVisualBasic.Completion
Imports Microsoft.SmallVisualBasic.Engine
Imports Microsoft.SmallVisualBasic.Statements

Namespace Microsoft.SmallVisualBasic.LanguageService
    Public Module CompilerService
        Private _compiler As Compiler
        Public Event HelpUpdated(itemWrapper As CompletionItemWrapper)


        Public Function GetTokens(
                     snapshot As ITextSnapshot,
                     ByRef startLine As Integer,
                     ByRef currentLine As Integer,
                     endLine As Integer) As List(Of Token)

            Dim tokens As New List(Of Token)
            Dim lineTokens As List(Of Token)
            Dim nextLineText = snapshot.GetLineFromLineNumber(startLine).GetText()
            Dim addNextLine = True

            For i = startLine To endLine - 1
                Dim curLineText = nextLineText
                nextLineText = snapshot.GetLineFromLineNumber(i + 1).GetText()
                lineTokens = LineScanner.GetTokens(curLineText, i - startLine)
                addNextLine = LineScanner.IsLineContinuity(lineTokens, nextLineText)

                If i = currentLine OrElse addNextLine Then
                    tokens.AddRange(lineTokens)
                Else
                    startLine = currentLine 'ignore prev line
                    If startLine = endLine Then
                        addNextLine = True
                        Exit For
                    End If
                End If
            Next

            If addNextLine Then
                lineTokens = LineScanner.GetTokens(nextLineText, endLine - startLine)
                tokens.AddRange(lineTokens)
            End If

            currentLine -= startLine  '  line pos abs for scaned tokens
            Return tokens
        End Function


        Public Function GetNextToken(ByRef i As Integer, direction As Integer, ByRef line As ITextSnapshotLine, ByRef tokens As List(Of Token)) As Token
            i += direction
            If i < 0 Then
                Dim lineNumber = line.LineNumber - 1
                If lineNumber < 0 Then Return Token.Illegal

                Dim snapshot = line.TextSnapshot
                line = snapshot.GetLineFromLineNumber(lineNumber)
                tokens = LineScanner.GetTokens(line.GetText(), lineNumber)
                i = tokens.Count - 1
                If i < 0 Then Return Token.Illegal

            ElseIf i >= tokens.Count Then
                Dim lineNumber = line.LineNumber + 1
                Dim snapshot = line.TextSnapshot
                If lineNumber >= snapshot.LineCount Then Return Token.Illegal

                line = snapshot.GetLineFromLineNumber(lineNumber)
                tokens = LineScanner.GetTokens(line.GetText(), lineNumber)
                If tokens.Count = 0 Then Return Token.Illegal
                i = 0
            End If

            Return tokens(i)

        End Function


        Public Sub ShowPopupHelp(itemWrapper As CompletionItemWrapper)
            RaiseEvent HelpUpdated(itemWrapper)
        End Sub

        Public ReadOnly Property DummyCompiler As Compiler
            Get

                If _compiler Is Nothing Then
                    _compiler = New Compiler()
                End If

                Return _compiler
            End Get
        End Property

        Public Function Compile(programText As String, errors As ICollection(Of String)) As Compiler
            Try
                Dim compiler As New Compiler()
                Dim list As List(Of [Error]) = compiler.Compile(New StringReader(programText))

                For Each item In list
                    errors.Add($"{item.Line + 1},{item.Column + 1}: {item.Description}")
                Next

                Return compiler
            Catch ex As Exception
                errors.Add(ex.Message)
                Return Nothing
            End Try
        End Function

        Public Sub FormatDocument(
                           textBuffer As ITextBuffer,
                           Optional lineNumber As Integer = -1,
                           Optional prettyListing As Boolean = True
                   )

            Try
                Dim snapshot = textBuffer.CurrentSnapshot
                Dim source As New TextBufferReader(snapshot)

                Dim indentationLevel = 0
                Dim start, [end] As Integer
                If lineNumber = -1 Then
                    start = 0
                    [end] = snapshot.LineCount - 1
                Else
                    start = FindCurrentSubStart(snapshot, lineNumber)
                    [end] = FindCurrentSubEnd(snapshot, lineNumber)
                End If

                Dim lines As New List(Of String)

                For i = start To [end]
                    lines.Add(snapshot.GetLineFromLineNumber(i).GetText())
                Next

                Using textEdit = textBuffer.CreateEdit()
                    Dim comments As New List(Of Token)

                    For lineNum = 0 To lines.Count - 1
                        Dim line = snapshot.GetLineFromLineNumber(lineNum + start)
                        ' (lineNum) to send it not to be changed ByRef
                        Dim tokens = LineScanner.GetTokens(lines(lineNum), (lineNum), lines)
                        comments.AddRange(LineScanner.SubLineComments)

                        If tokens.Count = 0 Then
                            AdjustIndentation(textEdit, line, indentationLevel, lines(lineNum).Length)
                            Continue For
                        End If

                        If prettyListing OrElse lineNum <> lineNumber Then
                            FormatLine(textEdit, line, tokens, 0)
                        End If

                        Dim firstCharPos = tokens(0).Column

                        Select Case tokens(0).Type
                            Case TokenType.EndIf, TokenType.Next, TokenType.EndFor, TokenType.Wend, TokenType.EndWhile
                                indentationLevel -= 1
                                AdjustIndentation(textEdit, line, indentationLevel, firstCharPos)

                            Case TokenType.EndSub, TokenType.EndFunction
                                indentationLevel = 0
                                AdjustIndentation(textEdit, line, indentationLevel, firstCharPos)

                            Case TokenType.If
                                AdjustIndentation(textEdit, line, indentationLevel, firstCharPos)
                                indentationLevel += 1
                                AddThen(snapshot, start, textEdit, lineNum, line, tokens)

                            Case TokenType.For, TokenType.ForEach, TokenType.While
                                AdjustIndentation(textEdit, line, indentationLevel, firstCharPos)
                                indentationLevel += 1

                            Case TokenType.Sub, TokenType.Function
                                AdjustIndentation(textEdit, line, 0, firstCharPos)
                                indentationLevel = 1

                            Case TokenType.Else
                                indentationLevel -= 1
                                AdjustIndentation(textEdit, line, indentationLevel, firstCharPos)
                                indentationLevel += 1
                                If tokens.Count > 1 AndAlso tokens(1).Type = TokenType.If Then
                                    textEdit.Replace(line.Start + tokens(0).Column, tokens(1).EndColumn - tokens(0).Column, "ElseIf")
                                    AddThen(snapshot, start, textEdit, lineNum, line, tokens)
                                End If

                            Case TokenType.ElseIf
                                indentationLevel -= 1
                                AdjustIndentation(textEdit, line, indentationLevel, firstCharPos)
                                indentationLevel += 1
                                AddThen(snapshot, start, textEdit, lineNum, line, tokens)

                            Case Else
                                If tokens.Count > 1 Then
                                    If tokens(0).LCaseText = "end" Then
                                        Select Case tokens(1).Type
                                            Case TokenType.If, TokenType.For, TokenType.While
                                                indentationLevel -= 1
                                                AdjustIndentation(textEdit, line, indentationLevel, firstCharPos)
                                                textEdit.Delete(line.Start + tokens(0).EndColumn, tokens(1).Column - tokens(0).EndColumn)

                                            Case TokenType.Sub, TokenType.Function
                                                indentationLevel = 0
                                                AdjustIndentation(textEdit, line, indentationLevel, firstCharPos)
                                                textEdit.Delete(line.Start + tokens(0).EndColumn, tokens(1).Column - tokens(0).EndColumn)

                                            Case Else
                                                AdjustIndentation(textEdit, line, indentationLevel, firstCharPos)
                                        End Select

                                    ElseIf tokens(0).Type <> TokenType.Question AndAlso tokens(1).Type = TokenType.Colon Then
                                        AdjustIndentation(textEdit, line, Integer.MaxValue, firstCharPos)
                                    Else
                                        AdjustIndentation(textEdit, line, indentationLevel, firstCharPos)
                                    End If
                                Else
                                    AdjustIndentation(textEdit, line, indentationLevel, firstCharPos)
                                End If
                        End Select

                        ' format sub lines
                        Dim lineStart = False, lineEnd = False, subLine = 0
                        Dim subLineOffset = 0
                        Dim indentStack As New Stack(Of Integer)
                        Dim firstSubLine = True
                        Dim n = tokens.Count - 1

                        For i = 1 To n
                            Dim t = tokens(i)
                            If t.subLine > subLine Then
                                lineStart = True
                                Do
                                    lineNum += 1
                                    line = snapshot.GetLineFromLineNumber(lineNum + start)
                                    If t.subLine - subLine > 1 Then
                                        ' An empty subline that contains only a hyphen and maybe a comment
                                        AdjustIndentation(textEdit, line, indentationLevel + subLineOffset, line.GetText().IndexOf("_"))
                                        Dim commentToken = GetSublineComment(lineNum)
                                        If Not commentToken.IsIllegal Then FormatComment(textEdit, line, commentToken)
                                        subLine += 1
                                    Else
                                        Exit Do
                                    End If
                                Loop
                                If prettyListing OrElse lineNum <> lineNumber Then FormatLine(textEdit, line, tokens, i)
                            Else
                                lineStart = False
                            End If

                            subLine = t.subLine
                            lineEnd = (i = n OrElse tokens(i + 1).subLine > subLine)

                            Select Case t.Type
                                Case TokenType.LeftParens, TokenType.LeftBracket, TokenType.LeftBrace
                                    If lineStart Then AdjustIndentation(textEdit, line, indentationLevel + subLineOffset, t.Column)
                                    If Not lineEnd Then
                                        subLineOffset += 1
                                        indentStack.Push(subLineOffset)
                                    End If

                                Case TokenType.RightParens, TokenType.RightBracket, TokenType.RightBrace
                                    If Not lineEnd Then
                                        If indentStack.Count = 0 Then
                                            subLineOffset = Math.Max(0, subLineOffset - 1)
                                        Else
                                            Dim l = indentStack.Pop()
                                            subLineOffset = If(indentStack.Count = 0, Math.Max(0, l - 1), indentStack.Peek())
                                        End If

                                        If lineStart Then AdjustIndentation(textEdit, line, indentationLevel + subLineOffset, t.Column)

                                    ElseIf lineStart Then
                                        subLineOffset = If(indentStack.Count = 0, Math.Max(0, subLineOffset - 1), Math.Max(0, indentStack.Peek() - 1))
                                        AdjustIndentation(textEdit, line, indentationLevel + subLineOffset, t.Column)
                                    Else
                                        subLineOffset = Math.Max(0, subLineOffset - 1)
                                    End If

                                Case TokenType.Concatenation, TokenType.Addition,
                                         TokenType.Subtraction, TokenType.Multiplication,
                                         TokenType.Division, TokenType.Mod,
                                         TokenType.Or, TokenType.And
                                    If lineStart Then
                                        subLineOffset = Math.Max(1, subLineOffset)
                                        AdjustIndentation(textEdit, line, indentationLevel + subLineOffset, t.Column)
                                    End If

                                Case Else
                                    If lineStart Then
                                        AdjustIndentation(textEdit, line, indentationLevel + subLineOffset, t.Column)
                                    End If
                            End Select

                            If Not lineEnd Then Continue For

                            'CheckLineEnd:
                            Dim last = If(tokens(i).ParseType = ParseType.Comment, i - 1, i)

                            If last = -1 Then
                                subLineOffset = 0
                                Continue For
                            End If

                            Select Case tokens(last).Type
                                Case TokenType.Comma
                                    If indentStack.Count = 0 Then
                                        If subLineOffset = 0 Then subLineOffset = 1
                                    Else
                                        subLineOffset = indentStack.Peek()
                                    End If

                                Case TokenType.Concatenation
                                    If n > 1 AndAlso tokens(last - 1).Type <> TokenType.Question Then
                                        subLineOffset = Math.Max(1, subLineOffset)
                                    End If

                                Case TokenType.Addition,
                                         TokenType.Subtraction, TokenType.Multiplication,
                                         TokenType.Division, TokenType.Mod,
                                         TokenType.And, TokenType.Or, TokenType.LineContinuity
                                    subLineOffset = Math.Max(1, subLineOffset)

                                Case TokenType.EqualsTo
                                    subLineOffset += 1

                                Case TokenType.LeftParens, TokenType.LeftBrace, TokenType.LeftBracket
                                    subLineOffset += 1
                                    indentStack.Push(subLineOffset)

                                Case TokenType.RightParens, TokenType.RightBrace, TokenType.RightBracket
                                    If indentStack.Count > 0 Then indentStack.Pop()

                                Case Else
                                    If tokens(last).Comment?.StartsWith("_") Then
                                        ' line ends with a hyphen
                                        subLineOffset += 1
                                    End If
                            End Select
                        Next
                    Next

                    textEdit.Apply()
                End Using

                FixKeywords(textBuffer, start, [end])
                FixIdentifiers(textBuffer, start, [end])

            Catch ex As Exception
            End Try
        End Sub

        Private Function GetSublineComment(lineNum As Integer) As Token
            For Each token In LineScanner.SubLineComments
                If token.Line = lineNum Then Return token
            Next
            Return Nothing
        End Function

        Private Sub FormatComment(
                     textEdit As ITextEdit,
                     line As ITextSnapshotLine,
                     token As Token)

            Dim comment = token.Text
            If comment.Length > 1 Then
                If Char.IsWhiteSpace(comment(1)) Then
                    Dim n = 0
                    For j = 2 To comment.Length - 1
                        If Char.IsWhiteSpace(comment(j)) Then
                            n += 1
                        Else
                            Exit For
                        End If
                    Next
                    If n > 0 Then textEdit.Replace(line.Start + token.Column + 2, n, "")
                Else
                    textEdit.Insert(line.Start + token.Column + 1, " ")
                End If
            End If
        End Sub

        Private Sub AddThen(
                           snapshot As ITextSnapshot,
                           start As Integer,
                           textEdit As ITextEdit,
                           lineNum As Integer,
                           line As ITextSnapshotLine,
                           tokens As List(Of Token))

            Dim n = tokens.Count
            If n > 1 Then
                Dim lastToken = tokens(n - 1)
                If lastToken.Type = TokenType.Comment Then
                    If n = 2 Then Return
                    lastToken = tokens(n - 2)
                End If

                If Not lastToken.IsIllegal AndAlso lastToken.Type <> TokenType.Then Then
                    Dim thenLine = If(lastToken.Line = lineNum, line, snapshot.GetLineFromLineNumber(lastToken.Line + start))
                    textEdit.Insert(thenLine.Start + lastToken.EndColumn, " Then")
                    tokens.Add(New Token() With {
                        .Column = lastToken.EndColumn + 1,
                        .Text = "Then",
                        .Line = lastToken.Line,
                        .subLine = lastToken.subLine,
                        .Type = TokenType.Then
                    })
                    If lineNum < snapshot.LineCount - 1 Then
                        Dim nextLine = snapshot.GetLineFromLineNumber(lineNum + 1)
                        Dim token = LineScanner.GetFirstToken(nextLine.GetText(), lineNum + 1)
                        If token.Type = TokenType.Then Then
                            textEdit.Delete(nextLine.Start + token.Column, token.EndColumn - token.Column)
                        End If
                    End If
                End If
            End If
        End Sub

        Private Function IsClosingSymbol(token As Token) As Boolean
            Select Case token.Type
                Case TokenType.RightParens, TokenType.RightBracket, TokenType.RightBrace
                    Return True
                Case Else
                    Return False
            End Select
        End Function

        Private Sub FixKeywords(textBuffer As ITextBuffer, start As Integer, [end] As Integer)
            Dim snapshot = textBuffer.CurrentSnapshot
            Dim lines As New List(Of String)

            For i = start To [end]
                lines.Add(snapshot.GetLineFromLineNumber(i).GetText())
            Next

            Using textEdit = textBuffer.CreateEdit()
                Dim moveCaret = False

                For lineNum = 0 To lines.Count - 1
                    Dim tokens = LineScanner.GetTokens(lines(lineNum), lineNum, lines)
                    Dim n = tokens.Count

                    For i = 0 To n - 1
                        Dim token = tokens(i)
                        Dim line = snapshot.GetLineFromLineNumber(token.Line + start)
                        If token.Type = TokenType.LineContinuity Then Continue For
                        If token.Type = TokenType.Colon Then Continue For

                        If token.Type = TokenType.Question Then
                            If i = 0 Then
                                If n = 1 Then
                                    textEdit.Replace(line.Start + token.Column, 1, "TW.WriteLine("""")")
                                ElseIf tokens(1).Type = TokenType.Colon AndAlso (n = 2 OrElse tokens(2).IsIllegal) Then
                                    textEdit.Replace(line.Start + token.Column, tokens(1).EndColumn - token.Column, "TW.Write("""")")
                                Else
                                    Dim lastToken = tokens(tokens.Count - 1)
                                    If lastToken.IsIllegal AndAlso lastToken.subLine > token.subLine Then
                                        If (tokens.Count < 3) Then Continue For
                                        lastToken = tokens(tokens.Count - 2)
                                    End If

                                    If lastToken.Type = TokenType.Comment Then
                                        If n < 3 Then Continue For
                                        lastToken = tokens(tokens.Count - 2)
                                    End If

                                    Dim length As Integer
                                    Dim method As String
                                    Dim closingChar = ")"

                                    Select Case tokens(1).Type
                                        Case TokenType.Colon
                                            length = tokens(2).Column - token.Column
                                            method = "TW.Write("
                                        Case TokenType.LeftBrace
                                            length = tokens(1).Column - token.Column
                                            method = "TW.WriteLines("
                                        Case Else
                                            length = tokens(1).Column - token.Column
                                            Dim args = Parser.ParseCommaSeparatedList(tokens, 1)
                                            If args.Count > 1 Then
                                                method = "TW.WriteLines({"
                                                closingChar = "})"
                                            Else
                                                method = "TW.WriteLine("
                                            End If
                                    End Select

                                    textEdit.Replace(line.Start + token.Column, length, method)
                                    ' Use actual last token
                                    lastToken = tokens(tokens.Count - 1)
                                    Dim lineStart = If(
                                         lastToken.subLine = 0,
                                         line.Start,
                                         snapshot.GetLineFromLineNumber(lastToken.Line + start).Start
                                     )

                                    textEdit.Insert(lineStart + lastToken.EndColumn, closingChar)
                                    If lastToken.IsIllegal AndAlso closingChar.Length = 2 Then
                                        moveCaret = True
                                    End If
                                End If

                            ElseIf tokens(i - 1).ParseType = ParseType.Operator OrElse
                                     tokens(i - 1).Type = TokenType.Question OrElse
                                     tokens(i - 1).Type = TokenType.To Then
                                textEdit.Replace(line.Start + token.Column, 1, "TW.Read()")
                            End If

                        ElseIf token.Type = TokenType.HashQuestion Then
                            If i = 0 OrElse tokens(i - 1).ParseType = ParseType.Operator OrElse
                                      tokens(i - 1).Type = TokenType.Question OrElse
                                      tokens(i - 1).Type = TokenType.To Then
                                Dim st = token.Column
                                textEdit.Replace(line.Start + st, token.EndColumn - st, "TW.ReadNumber()")
                            End If

                        ElseIf i = 0 AndAlso token.Type = TokenType.Identifier AndAlso token.LCaseText = "msgbox" Then
                            Dim symbolTable = GetSymbolTable(textBuffer)
                            If symbolTable.Subroutines.ContainsKey("msgbox") Then Continue For

                            Dim length = 0
                            Dim startsWithLeftParens = False
                            Dim needsClosing = TokenCount(tokens, TokenType.LeftParens) > TokenCount(tokens, TokenType.RightParens)
                            ' Use actual last token
                            Dim lastToken = tokens(tokens.Count - 1)
                            If tokens.Count = 1 Then
                                length = token.EndColumn - token.Column
                            ElseIf tokens(1).Type = TokenType.LeftParens AndAlso (
                                    needsClosing OrElse lastToken.Type = TokenType.RightParens) Then
                                length = If(tokens.Count > 2, tokens(2).Column, tokens(1).EndColumn) - token.Column
                                startsWithLeftParens = True
                            Else
                                length = tokens(1).Column - token.Column
                            End If

                            textEdit.Replace(line.Start + token.Column, length, "Forms.ShowMessage(")
                            Dim lineStart = If(
                                lastToken.subLine = 0,
                                line.Start,
                                snapshot.GetLineFromLineNumber(lastToken.Line + start).Start
                            )
                            Dim args = Parser.ParseCommaSeparatedList(tokens, If(startsWithLeftParens, 2, 1))
                            If args.Count = 0 Then
                                If needsClosing Then
                                    textEdit.Insert(lineStart + lastToken.EndColumn, """"", ""Message"")")
                                Else
                                    textEdit.Insert(lineStart + lastToken.Column, """"", ""Message""")
                                End If
                            ElseIf args.Count > 1 Then
                                    If Not startsWithLeftParens OrElse lastToken.Type <> TokenType.RightParens OrElse
                                        needsClosing Then
                                        textEdit.Insert(lineStart + lastToken.EndColumn, ")")
                                    End If
                                ElseIf startsWithLeftParens AndAlso
                                        lastToken.Type = TokenType.RightParens AndAlso
                                        TokenCount(tokens, TokenType.LeftParens) = TokenCount(tokens, TokenType.RightParens) Then
                                    textEdit.Insert(lineStart + lastToken.Column, ", ""Message""")
                                Else
                                    textEdit.Insert(lineStart + lastToken.EndColumn, ", ""Message"")")
                            End If

                        ElseIf token.ParseType = ParseType.Keyword OrElse
                                        token.Type = TokenType.And OrElse
                                        token.Type = TokenType.Or OrElse
                                        token.Type = TokenType.Mod Then
                            Dim keyword = token.Type.ToString()
                            If token.Text <> keyword Then
                                textEdit.Replace(line.Start + token.Column, token.EndColumn - token.Column, keyword)
                            End If
                        End If
                    Next
                Next

                textEdit.Apply()
                If moveCaret Then
                    Dim document = textBuffer.Properties.GetProperty("Document")
                    If document IsNot Nothing Then
                        Dim textView = CType(document.EditorControl.TextView, Editor.ITextView)
                        textView.Caret.MoveTo(textView.Caret.Position.TextInsertionIndex - 2)
                    End If
                End If
            End Using
        End Sub

        Private Function TokenCount(tokens As List(Of Token), type As TokenType) As Object
            Dim n = 0
            For Each token In tokens
                If token.Type = type Then n += 1
            Next
            Return n
        End Function

        Private Sub FixIdentifiers(
                            textBuffer As ITextBuffer,
                            start As Integer,
                            [end] As Integer
                    )

            Dim snapshot = textBuffer.CurrentSnapshot
            Dim symbolTable = GetSymbolTable(textBuffer)

            Using textEdit = textBuffer.CreateEdit()
                ' fix lib types and members
                For Each token In symbolTable.AllLibMembers
                    Dim line = snapshot.GetLineFromLineNumber(token.Line)
                    If line.LineNumber < start OrElse line.LineNumber > [end] Then Continue For

                    ' The exact type/method name is stored in the comment field
                    textEdit.Replace(line.Start + token.Column, token.EndColumn - token.Column, token.Comment)
                Next

                ' fix local vars definitions
                Dim locals = symbolTable.LocalVariables
                For i = 0 To locals.Count - 1
                    Dim expr = locals.Values(i)
                    Dim id = expr.Identifier
                    Dim idName = id.Text
                    Dim n = 0

                    If idName.StartsWith("_") Then
                        If idName.Length = 1 Then Continue For
                        n = 1
                    End If

                    If Char.IsUpper(idName(n)) Then
                        Dim line = snapshot.GetLineFromLineNumber(id.Line)
                        Dim c = idName(n).ToString().ToLower()
                        textEdit.Replace(line.Start + id.Column + n, 1, c)
                        id.Text = idName.Substring(0, n) & c + If(idName.Length > n + 1, idName.Substring(n + 1), "")
                        expr.Identifier = id
                        symbolTable.LocalVariables(locals.Keys(i)) = expr
                    End If
                Next

                ' fix global vars definitions
                ToUpper(start, [end], snapshot, symbolTable.GlobalVariables, textEdit)

                ' fix subs definitions
                ToUpper(start, [end], snapshot, symbolTable.Subroutines, textEdit)

                ' fix labels definitions
                ToUpper(start, [end], snapshot, symbolTable.Labels, textEdit)

                ' fix dynamic properties definitions
                For Each dynObj In symbolTable.Dynamics
                    ToUpper(start, [end], snapshot, dynObj.Value, textEdit)
                Next

                Dim controlNames = symbolTable.ControlNames

                ' fix vars usage
                For Each id In symbolTable.AllIdentifiers
                    Dim line = snapshot.GetLineFromLineNumber(id.Line)
                    If line.LineNumber < start OrElse line.LineNumber > [end] Then Continue For

                    If id.LCaseText = "me" Then
                        If id.Text <> "Me" Then textEdit.Replace(line.Start + id.Column, id.EndColumn - id.Column, "Me")
                        FixParans(id, line, textEdit, True)
                        Continue For
                    End If

                    Select Case id.SymbolType
                        Case CompletionItemType.LocalVariable
                            ' fix local vars usage
                            Dim subName = Statements.SubroutineStatement.GetSubroutine(id)?.Name.LCaseText
                            Dim key = $"{subName}.{id.LCaseText}"
                            Dim definition = symbolTable.LocalVariables(key).Identifier
                            If id.Line <> definition.Line OrElse id.Column <> definition.Column Then
                                Dim name = definition.Text
                                If id.Text <> name Then textEdit.Replace(line.Start + id.Column, id.EndColumn - id.Column, name)
                            End If
                            FixParans(id, line, textEdit, True)

                        Case CompletionItemType.GlobalVariable
                            ' fix global vars usage
                            If FixToken(id, line.Start, symbolTable.GlobalVariables, textEdit) Then
                                FixParans(id, line, textEdit, True)
                                Continue For
                            End If

                            Dim fixed = False
                            If controlNames IsNot Nothing Then
                                ' fix control names usage
                                Dim controlId = id.LCaseText
                                For Each controlName In controlNames
                                    If controlName.ToLower() = controlId Then
                                        If id.Text <> controlName Then textEdit.Replace(line.Start + id.Column, id.EndColumn - id.Column, controlName)
                                        fixed = True
                                        FixParans(id, line, textEdit, True)
                                        Exit For
                                    End If
                                Next
                            End If

                            ' Add missing parans to subroutine calls
                            If Not fixed AndAlso FixToken(id, line.Start, symbolTable.Subroutines, textEdit) Then
                                FixParans(id, line, textEdit)
                            End If

                        Case CompletionItemType.SubroutineName
                            ' fix event handlers and sub call
                            If TypeOf id.Parent Is SubroutineStatement Then
                                FixParans(id, line, textEdit)
                            ElseIf FixToken(id, line.Start, symbolTable.Subroutines, textEdit) Then
                                If IsEventHandler(id) Then FixParans(id, line, textEdit, True)
                            Else
                                FixParans(id, line, textEdit, True)
                            End If

                        Case CompletionItemType.Label
                            ' fix goto labels
                            FixToken(id, line.Start, symbolTable.Labels, textEdit)
                            FixParans(id, line, textEdit, True)

                        Case CompletionItemType.PropertyName
                            Dim tokens = LineScanner.GetTokens(line.GetText(), id.Line)
                            Dim n = tokens.Count - 1
                            For i = 0 To n
                                If tokens(i).Column = id.Column Then
                                    If i - 2 >= 0 Then
                                        Dim m = symbolTable.GetMethodInfo(tokens(i - 2), id)
                                        If m IsNot Nothing Then
                                            If i = n OrElse tokens(i + 1).Type <> TokenType.LeftParens Then
                                                textEdit.Insert(line.Start + id.EndColumn, "()")
                                            End If
                                        End If
                                    End If
                                    Exit For
                                End If
                            Next

                        Case CompletionItemType.MethodName
                            Dim tokens = LineScanner.GetTokens(line.GetText(), id.Line)
                            Dim n = tokens.Count - 1
                            For i = 0 To n
                                If tokens(i).Column = id.Column Then
                                    If i - 2 >= 0 Then
                                        Dim p = symbolTable.GetPropertyInfo(tokens(i - 2), id)
                                        If p IsNot Nothing Then
                                            If i < n - 1 AndAlso tokens(i + 1).Type = TokenType.LeftParens AndAlso tokens(i + 2).Type = TokenType.RightParens Then
                                                textEdit.Replace(line.Start + id.EndColumn, tokens(i + 2).EndColumn - id.EndColumn, "")
                                            End If
                                        End If
                                    End If
                                    Exit For
                                End If
                            Next
                    End Select
                Next

                For Each obj In symbolTable.AllDynamicProperties
                    If Not symbolTable.Dynamics.ContainsKey(obj.Key) Then Continue For

                    Dim objName = obj.Key
                    Dim x = CompletionHelper.TrimData(objName)
                    Dim propDictionery = symbolTable.Dynamics(objName)
                    Dim fixed = False

                    For Each prop In obj.Value
                        Dim line = snapshot.GetLineFromLineNumber(prop.Line)
                        If line.LineNumber < start OrElse line.LineNumber > [end] Then Continue For

                        If Not FixToken(prop, line.Start, propDictionery, textEdit) Then
                            For Each type In symbolTable.Dynamics
                                If type.Key = objName Then Continue For ' Add before
                                Dim y = CompletionHelper.TrimData(type.Key)
                                If x.Contains(y) Then
                                    Dim propDictionery2 = symbolTable.Dynamics(type.Key)
                                    If FixToken(prop, line.Start, propDictionery2, textEdit) Then
                                        FixParans(prop, line, textEdit, True)
                                        Exit For
                                    End If
                                End If
                            Next
                        End If
                    Next
                Next

                textEdit.Apply()
            End Using
        End Sub

        Private Function IsEventHandler(id As Token) As Boolean
            Dim assignment = TryCast(id.Parent, AssignmentStatement)
            If assignment Is Nothing Then Return False
            Dim prop = TryCast(assignment.LeftValue, Expressions.PropertyExpression)
            If prop Is Nothing Then Return False
            Return prop.IsEvent
        End Function

        Private Sub FixParans(
                           id As Token,
                           line As ITextSnapshotLine,
                           textEdit As ITextEdit,
                           Optional RemoveParans As Boolean = False)

            Dim tokens = LineScanner.GetTokens(line.GetText(), id.Line)
            Dim n = tokens.Count - 1
            For i = 0 To n
                If tokens(i).Column = id.Column Then
                    If RemoveParans Then
                        If i < n - 1 AndAlso tokens(i + 1).Type = TokenType.LeftParens AndAlso tokens(i + 2).Type = TokenType.RightParens Then
                            textEdit.Replace(line.Start + id.EndColumn, tokens(i + 2).EndColumn - id.EndColumn, "")
                        End If
                    ElseIf i = n OrElse tokens(i + 1).Type <> TokenType.LeftParens Then
                        textEdit.Insert(line.Start + id.EndColumn, "()")
                        Return
                    End If
                End If
            Next
        End Sub

        Private Function FixToken(
                            token As Token,
                            lineStart As Integer,
                            dictionary As Dictionary(Of String, Token),
                            textEdit As ITextEdit) As Boolean

            Dim key = token.LCaseText
            If key.StartsWith("__foreach__") Then Return True

            If dictionary.ContainsKey(key) Then
                Dim defenition = dictionary(key)
                If defenition.Line <> token.Line OrElse defenition.Column <> token.Column Then
                    Dim name = defenition.Text
                    If token.Text <> name Then textEdit.Replace(lineStart + token.Column, token.EndColumn - token.Column, name)
                End If
                Return True
            End If

            Return False
        End Function

        Private Sub ToUpper(start As Integer, [end] As Integer, snapshot As ITextSnapshot, dictionary As Dictionary(Of String, Token), textEdit As ITextEdit)
            For i = 0 To dictionary.Count - 1
                Dim token = dictionary.Values(i)
                Dim name = token.Text
                Dim n = 0

                If name.StartsWith("_") Then
                    If name.Length = 1 Then Continue For
                    n = 1
                End If

                If Char.IsLower(name(n)) Then
                    Dim line = snapshot.GetLineFromLineNumber(token.Line)
                    If line.LineNumber >= start AndAlso line.LineNumber <= [end] Then
                        Dim c = token.Text(n).ToString().ToUpper()
                        textEdit.Replace(line.Start + token.Column + n, 1, c)
                        token.Text = name.Substring(0, n) & c + If(name.Length > n + 1, name.Substring(n + 1), "")
                        dictionary(dictionary.Keys(i)) = token
                    End If
                End If
            Next
        End Sub

        Public Function GetSymbolTable(buffer As ITextBuffer) As SymbolTable
            Dim compiler = DummyCompiler ' important to force the property to create the compilor
            Dim symbolTable = compiler.Parser.SymbolTable
            symbolTable.ModuleNames = buffer.Properties.GetProperty(Of Dictionary(Of String, String))("ControlsInfo")
            symbolTable.ControlNames = GetControlNames(buffer)
            compiler.Compile(New TextBufferReader(buffer.CurrentSnapshot), True)
            Return symbolTable
        End Function

        Friend Function GetControlNames(buffer As ITextBuffer) As List(Of String)
            Dim controlNames = buffer.Properties.GetProperty(Of List(Of String))("ControlNames")
            If controlNames Is Nothing Then Return New List(Of String)

            Return (From name In controlNames
                    Where name.ToLower() <> "graphicswindow").ToList()
        End Function

        Public Function FindCurrentSubStart(textSnapshot As ITextSnapshot, LineNumber As Integer) As Integer
            For i = LineNumber To 0 Step -1
                Dim line = textSnapshot.GetLineFromLineNumber(i)
                Dim token = LineScanner.GetFirstToken(line.GetText(), i)

                Select Case token.Type
                    Case TokenType.Sub, TokenType.Function
                        Return Math.Max(0, i - 1)  ' i -1 To format prev EndSub in special case

                    Case TokenType.EndSub, TokenType.EndFunction
                        If i < LineNumber Then
                            Return i  'not i + 1 To format prev End Sub
                        End If

                End Select
            Next

            Return 0
        End Function

        Public Function FindCurrentSubEnd(textSnapshot As ITextSnapshot, LineNumber As Integer) As Integer
            For i = LineNumber + 1 To textSnapshot.LineCount - 1
                Dim line = textSnapshot.GetLineFromLineNumber(i)
                Dim token = LineScanner.GetFirstToken(line.GetText(), i)
                Select Case token.Type
                    Case TokenType.Sub, TokenType.Function
                        If i > LineNumber + 1 Then Return i - 1

                    Case TokenType.EndSub, TokenType.EndFunction
                        Return i
                End Select
            Next

            Return textSnapshot.LineCount - 1
        End Function

        Private Sub FormatLine(
                         textEdit As ITextEdit,
                         line As ITextSnapshotLine,
                         tokens As List(Of Token),
                         startAt As Integer
                    )

            Dim text = line.GetText()
            Dim textLen = text.Length
            Dim endAt = tokens.Count - 1

            Dim subLine = tokens(startAt).subLine

            For i = startAt To endAt
                Dim token = tokens(i)
                If token.subLine <> subLine Then
                    token = GetSublineComment(tokens(i - 1).Line)
                    If Not token.IsIllegal Then
                        FixSpaces(textEdit, line, tokens(i - 1), token, 1)
                        FormatComment(textEdit, line, token)
                    End If
                    Return
                End If

                Dim nextToken As Token
                Dim notLastToken = False
                If i < endAt Then
                    nextToken = tokens(i + 1)
                    notLastToken = nextToken.subLine = subLine
                End If

                Dim FixLiterals =
                    Sub()
                        If notLastToken Then
                            Select Case nextToken.Type
                                Case TokenType.Dot, TokenType.Lookup, TokenType.Comma, TokenType.Colon,
                                         TokenType.LeftBracket, TokenType.LeftBrace, TokenType.LeftParens,
                                         TokenType.RightBracket, TokenType.RightBrace, TokenType.RightParens
                                    FixSpaces(textEdit, line, token, nextToken, 0)
                                Case Else
                                    FixSpaces(textEdit, line, token, nextToken, 1)
                            End Select
                        End If
                    End Sub

                Select Case token.Type
                    Case TokenType.EqualsTo, TokenType.NotEqualsTo,
                             TokenType.GreaterThanOrEqualsTo,
                             TokenType.LessThanOrEqualsTo,
                             TokenType.And, TokenType.Or,
                             TokenType.Concatenation, TokenType.Multiplication,
                             TokenType.Division, TokenType.Mod
                        If notLastToken Then FixSpaces(textEdit, line, token, nextToken, 1)

                    Case TokenType.Addition
                        ' Remove unary + sign, because the compiler doeesn't recognize it
                        Dim deleted = False
                        If i = 0 Then
                            textEdit.Delete(line.Start + token.Column, 1)
                            deleted = True
                        Else
                            Select Case tokens(i - 1).Type
                                Case TokenType.Identifier,
                                         TokenType.DateLiteral, TokenType.NumericLiteral, TokenType.StringLiteral,
                                         TokenType.RightBrace, TokenType.RightBracket, TokenType.RightParens,
                                         TokenType.HashQuestion
                                    ' Additon allowed

                                Case TokenType.Question
                                    If i = 1 Then ' WriteLine( + )
                                        textEdit.Delete(line.Start + token.Column, 1)
                                        deleted = True
                                    End If

                                Case TokenType.Colon
                                    If i = 2 AndAlso tokens(0).Type = TokenType.Question Then ' Write( + )
                                        textEdit.Delete(line.Start + token.Column, 1)
                                        deleted = True
                                    End If

                                Case Else
                                    textEdit.Delete(line.Start + token.Column, 1)
                                    deleted = True
                            End Select
                        End If

                        If Not deleted AndAlso notLastToken Then
                            FixSpaces(textEdit, line, token, nextToken, 1)
                        End If

                    Case TokenType.LessThan
                        If notLastToken Then
                            If nextToken.Type = TokenType.GreaterThan OrElse nextToken.Type = TokenType.EqualsTo Then
                                FixSpaces(textEdit, line, token, nextToken, 0)
                            Else
                                FixSpaces(textEdit, line, token, nextToken, 1)
                            End If
                        End If

                    Case TokenType.GreaterThan
                        If notLastToken Then
                            If nextToken.Type = TokenType.EqualsTo Then
                                FixSpaces(textEdit, line, token, nextToken, 0)
                            Else
                                FixSpaces(textEdit, line, token, nextToken, 1)
                            End If
                        End If

                    Case TokenType.Subtraction
                        If notLastToken Then
                            Dim prevType = If(i = 0, ParseType.Illegal, tokens(i - 1).ParseType)
                            If i = 0 Then
                                FixSpaces(textEdit, line, token, nextToken, 0)
                            ElseIf prevType = ParseType.Operator OrElse prevType = ParseType.Keyword Then
                                Select Case tokens(i - 1).Type
                                    Case TokenType.RightBracket, TokenType.RightBrace, TokenType.RightParens
                                        FixSpaces(textEdit, line, token, nextToken, 1)
                                    Case Else
                                        FixSpaces(textEdit, line, token, nextToken, 0)
                                End Select
                            Else
                                FixSpaces(textEdit, line, token, nextToken, 1)
                            End If
                        End If

                    Case TokenType.LeftBracket, TokenType.LeftBrace, TokenType.LeftParens
                        If notLastToken Then FixSpaces(textEdit, line, token, nextToken, 0)

                    Case TokenType.RightBracket, TokenType.RightBrace, TokenType.RightParens,
                             TokenType.Question, TokenType.HashQuestion
                        If notLastToken Then
                            Select Case nextToken.Type
                                Case TokenType.Comma, TokenType.Dot,
                                        TokenType.LeftBracket, TokenType.RightBracket,
                                        TokenType.LeftBrace, TokenType.RightBrace,
                                        TokenType.LeftParens, TokenType.RightParens
                                    FixSpaces(textEdit, line, token, nextToken, 0)
                                Case Else
                                    FixSpaces(textEdit, line, token, nextToken, 1)
                            End Select
                        End If

                    Case TokenType.Comma
                        If notLastToken Then FixSpaces(textEdit, line, token, nextToken, 1)

                    Case TokenType.NumericLiteral
                        Dim t = token.Text

                        If t.EndsWith(".") Then
                            textEdit.Insert(line.Start + token.EndColumn, "0")
                        End If

                        FixLiterals()

                        If t.StartsWith(".") Then
                            textEdit.Insert(line.Start + token.Column, "0")
                        End If

                    Case TokenType.DateLiteral
                        FixLiterals()

                        Dim d = token.Text
                        Dim length = d.Length
                        If length > 1 Then
                            Dim index = 1

                            Do
                                If d(index) <> " " Then Exit Do
                                index += 1
                            Loop While index < length

                            If index > 1 Then textEdit.Replace(line.Start + token.Column + 1, index - 1, "")

                            index = length - 2
                            Do
                                If d(index) <> " " Then Exit Do
                                index -= 1
                            Loop While index > 0

                            If index < length - 2 Then textEdit.Replace(line.Start + token.Column + index + 1, length - 2 - index, "")
                        End If

                    Case TokenType.Identifier, TokenType.StringLiteral,
                             TokenType.True, TokenType.False
                        FixLiterals()

                    Case TokenType.Dot, TokenType.Lookup
                        If notLastToken Then FixSpaces(textEdit, line, token, nextToken, 0)

                    Case TokenType.Comment
                        FormatComment(textEdit, line, token)

                    Case Else
                        If token.ParseType = ParseType.Keyword Then
                            If notLastToken Then
                                If token.Type = TokenType.For AndAlso nextToken.LCaseText = "each" Then
                                    FixSpaces(textEdit, line, token, nextToken, 0)
                                Else
                                    FixSpaces(textEdit, line, token, nextToken, 1)
                                End If
                            End If
                        End If
                End Select

                If Not notLastToken Then
                    If token.Comment?.StartsWith("_") Then
                        Dim column = CInt(token.Comment.Substring(1))
                        If column = token.EndColumn Then
                            textEdit.Insert(line.Start + token.EndColumn, " ")
                        End If
                    End If
                End If
            Next

        End Sub

        Private Sub FixSpaces(
                       textEdit As ITextEdit,
                       line As ITextSnapshotLine,
                       token1 As Token,
                       token2 As Token,
                       requiredSpaces As Integer
                )

            If token2.Type = TokenType.Comment AndAlso
                token2.Column > token1.EndColumn Then Return

            Dim spaces = token2.Column - token1.EndColumn
            If spaces < 0 OrElse spaces = requiredSpaces Then Return

            Dim n = spaces - requiredSpaces
            If n > 0 Then
                textEdit.Replace(line.Start + token2.Column - n, n, "")
            Else
                textEdit.Insert(line.Start + token2.Column, Space(n * -1))
            End If
        End Sub

        Private Sub AdjustIndentation(
                    textEdit As ITextEdit,
                    line As ITextSnapshotLine,
                    indentationLevel As Integer,
                    firstCharPos As Integer
                )

            Dim indent As Integer
            If indentationLevel < 0 Then
                indent = 0
            ElseIf indentationLevel = Integer.MaxValue Then ' Label
                indent = 1
            Else
                indent = indentationLevel * 3
            End If

            If firstCharPos <> indent Then
                textEdit.Replace(
                    line.Start, firstCharPos,
                    New String(" "c, indent)
                )
            End If

            ' Trim line end
            Dim x = line.GetText()
            Dim L = x.Length - x.TrimEnd().Length
            If L > 0 AndAlso L < x.Length Then
                Select Case x.Trim().ToLower()
                    Case "sub", "function", "while"
                    Case Else
                        textEdit.Delete(New Span(line.Start + line.Length - L, L))
                End Select
            End If
        End Sub

        Public Function GetLastTokenIndex(line As ITextSnapshotLine, caretPosition As Integer, tokens As List(Of Token)) As Integer
            Dim lineNumber = line.LineNumber
            Dim column = caretPosition - line.Start

            Dim n As Integer
            For n = 0 To tokens.Count - 1
                Dim token = tokens(n)
                Dim tokenLine = token.Line
                If tokenLine > lineNumber Then Exit For
                If tokenLine = lineNumber AndAlso (
                        token.Column > column OrElse
                        (token.Type = TokenType.StringLiteral AndAlso token.EndColumn = column AndAlso token.EndColumn - token.Column = 1)
                    ) Then Exit For
            Next

            Return n - 1
        End Function

    End Module
End Namespace
