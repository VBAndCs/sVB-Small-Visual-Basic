Imports System
Imports System.Collections.Generic
Imports System.IO
Imports Microsoft.Nautilus.Text
Imports Microsoft.Nautilus.Text.Editor
Imports Microsoft.SmallBasic.Completion

Namespace Microsoft.SmallBasic.LanguageService
    Public Module CompilerService
        Private _compiler As Compiler

        Public ReadOnly Property DummyCompiler As Compiler
            Get

                If _compiler Is Nothing Then
                    _compiler = New Compiler()
                End If

                Return _compiler
            End Get
        End Property

        Public Event CurrentCompletionItemChanged As EventHandler(Of CurrentCompletionItemChangedEventArgs)

        Public Function Compile(programText As String, outputFilePath As String, errors As ICollection(Of String)) As Boolean
            Try
                Dim compiler As New Compiler()
                Dim fileNameWithoutExtension = Path.GetFileNameWithoutExtension(outputFilePath)
                Dim directoryName = Path.GetDirectoryName(outputFilePath)
                Dim list As List(Of [Error]) = compiler.Build(New StringReader(programText), fileNameWithoutExtension, directoryName)

                For Each item In list
                    errors.Add($"{item.Line + 1},{item.Column + 1}: {item.Description}")
                Next

                Return errors.Count = 0
            Catch ex As Exception
                errors.Add(ex.Message)
                Return False
            End Try
        End Function

        Public Function Compile(programText As String, errors As ICollection(Of String)) As Compiler
            Try
                Dim compiler As Compiler = New Compiler()
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

        Public Sub UpdateCurrentCompletionItem(ByVal completionItemWrapper As CompletionItemWrapper)
            RaiseEvent CurrentCompletionItemChanged(Nothing, New CurrentCompletionItemChangedEventArgs(completionItemWrapper))
        End Sub

        Public Sub FormatDocument(textBuffer As ITextBuffer, Optional lineNumber As Integer = -1, Optional prettyListing As Boolean = True)
            Dim currentSnapshot = textBuffer.CurrentSnapshot
            Dim source As New TextBufferReader(currentSnapshot)
            Dim completionHelper As New CompletionHelper()
            Dim textEdit = textBuffer.CreateEdit()

            Using textEdit
                Dim indentationLevel = 0
                Dim start, [end] As Integer
                If lineNumber = -1 Then
                    start = 0
                    [end] = currentSnapshot.LineCount - 1
                Else
                    start = FindCurrentSubStart(currentSnapshot, lineNumber)
                    [end] = FindCurrentSubEnd(currentSnapshot, lineNumber)
                End If

                Dim lines As New List(Of String)

                For i = start To [end]
                    lines.Add(currentSnapshot.Lines(i).GetText())
                Next

                For lineNum = 0 To lines.Count - 1
                    Dim line = currentSnapshot.Lines(lineNum + start)
                    ' (lineNum) to send it ByVal not to be changed ByRef
                    Dim tokens = LineScanner.GetTokens(lines(lineNum), (lineNum), lines)

                    If tokens.Count = 0 Then
                        AdjustIndentation(textEdit, line, indentationLevel, lines(lineNum).Length)
                        Continue For
                    End If

                    If prettyListing OrElse lineNum <> lineNumber Then FormatLine(textEdit, line, tokens, 0)

                    Dim firstCharPos = tokens(0).Column

                    Select Case tokens(0).Token
                        Case Token.EndIf, Token.Next, Token.EndFor, Token.Wend, Token.EndWhile
                            indentationLevel -= 1
                            AdjustIndentation(textEdit, line, indentationLevel, firstCharPos)

                        Case Token.EndSub, Token.EndFunction
                            indentationLevel = 0
                            AdjustIndentation(textEdit, line, indentationLevel, firstCharPos)

                        Case Token.If, Token.For, Token.While
                            AdjustIndentation(textEdit, line, indentationLevel, firstCharPos)
                            indentationLevel += 1

                        Case Token.Sub, Token.Function
                            AdjustIndentation(textEdit, line, 0, firstCharPos)
                            indentationLevel = 1

                        Case Token.Else, Token.ElseIf
                            indentationLevel -= 1
                            AdjustIndentation(textEdit, line, indentationLevel, firstCharPos)
                            indentationLevel += 1

                        Case Else
                            AdjustIndentation(textEdit, line, indentationLevel, firstCharPos)
                    End Select

                    ' format sub lines
                    Dim lineStart = False, lineEnd = False, subLine = 0

                    Dim subLineOffset = 0
                    Dim indentStack As New Stack(Of Integer)
                    Dim firstSubLine = True
                    Dim n = tokens.Count - 1

                    For i = 1 To n
                        Dim t = tokens(i)
                        If t.subLine = 0 Then Continue For

                        If firstSubLine Then
                            firstSubLine = False
                            ' check the last token in the first line
                            i = i - 1
                            GoTo CheckLineEnd
                        End If

                        If t.subLine > subLine Then
                            lineStart = True
                            lineNum += 1
                            line = currentSnapshot.Lines(lineNum + start)
                            If prettyListing OrElse lineNum <> lineNumber Then FormatLine(textEdit, line, tokens, i)
                        Else
                            lineStart = False
                        End If

                        subLine = t.subLine
                        lineEnd = (i = n OrElse tokens(i + 1).subLine > subLine)

                        Select Case t.Token
                            Case Token.LeftParens, Token.LeftBracket, Token.LeftCurlyBracket
                                If lineStart Then AdjustIndentation(textEdit, line, indentationLevel + subLineOffset, t.Column)
                                If Not lineEnd Then
                                    subLineOffset += 1
                                    indentStack.Push(subLineOffset)
                                End If

                            Case Token.RightParens, Token.RightBracket, Token.RightCurlyBracket
                                If Not lineEnd Then
                                    If indentStack.Count = 0 Then
                                        subLineOffset = Math.Max(0, subLineOffset - 1)
                                    Else
                                        indentStack.Pop()
                                        subLineOffset = If(indentStack.Count = 0, Math.Max(0, subLineOffset - 1), indentStack.Peek())
                                    End If

                                ElseIf lineStart Then
                                    subLineOffset = If(indentStack.Count = 0, Math.Max(0, subLineOffset - 1), Math.Max(0, indentStack.Peek() - 1))
                                End If

                                If lineStart Then AdjustIndentation(textEdit, line, indentationLevel + subLineOffset, t.Column)

                            Case Token.Addition, Token.Subtraction, Token.Multiplication, Token.Division, Token.Or, Token.And
                                If lineStart Then
                                    subLineOffset = Math.Max(1, subLineOffset)
                                    AdjustIndentation(textEdit, line, indentationLevel + subLineOffset, t.Column)
                                End If

                            Case Else
                                If lineStart Then AdjustIndentation(textEdit, line, indentationLevel + subLineOffset, t.Column)
                        End Select

                        If Not lineEnd Then Continue For

CheckLineEnd:
                        Dim last = If(tokens(i).TokenType = TokenType.Comment, i - 1, i)

                        If last = -1 Then
                            subLineOffset = 0
                            Continue For
                        End If

                        Select Case tokens(last).NormalizedText
                            Case ","
                                If indentStack.Count = 0 Then
                                    If subLineOffset = 0 Then subLineOffset = 1
                                Else
                                    subLineOffset = indentStack.Peek()
                                End If

                            Case "_", "+", "-", "*", "/", "and", "or"
                                subLineOffset = Math.Max(1, subLineOffset)
                            Case "="
                                subLineOffset += 1
                            Case "(", "{", "["
                                subLineOffset += 1
                                indentStack.Push(subLineOffset)

                            Case ")", "}", "]"
                                If indentStack.Count > 0 Then indentStack.Pop()
                        End Select

                    Next

                Next

                textEdit.Apply()
            End Using
        End Sub

        Public Function FindCurrentSubStart(text As ITextSnapshot, LineNumber As Integer) As Integer
            For i = LineNumber To 0 Step -1
                Dim line = text.GetLineFromLineNumber(i)
                Dim tokenInfo = LineScanner.GetFirstToken(line.GetText(), i)

                Select Case tokenInfo.Token
                    Case Token.Sub, Token.Function
                        Return Math.Max(0, i - 1)  ' i -1 To format prev EndSub in special case

                    Case Token.EndSub, Token.EndFunction
                        If i < LineNumber Then
                            Return i  'not i + 1 To format prev End Sub
                        End If

                End Select
            Next

            Return 0
        End Function

        Public Function FindCurrentSubEnd(text As ITextSnapshot, LineNumber As Integer) As Integer
            For i = LineNumber + 1 To text.LineCount - 1
                Dim line = text.GetLineFromLineNumber(i)
                Dim tokenInfo = LineScanner.GetFirstToken(line.GetText(), i)
                Select Case tokenInfo.Token
                    Case Token.Sub, Token.Function
                        If i > LineNumber + 1 Then Return i - 1

                    Case Token.EndSub, Token.EndFunction
                        Return i
                End Select
            Next

            Return text.LineCount - 1
        End Function


        Private Sub FormatLine(
                         textEdit As ITextEdit,
                         line As ITextSnapshotLine,
                         tokens As List(Of TokenInfo),
                         startAt As Integer
                    )

            Dim text = line.GetText()
            Dim textLen = text.Length
            Dim endAt = tokens.Count - 1

            Dim subLine = tokens(startAt).subLine

            For i = startAt To endAt
                Dim t = tokens(i)
                If t.subLine <> subLine Then Return

                Dim notLastToken = i < endAt AndAlso tokens(i + 1).subLine = subLine

                Select Case t.Token
                    Case Token.Equals, Token.GreaterThan, Token.GreaterThanEqualTo, Token.LessThan, Token.LessThanEqualTo,
                                   Token.And, Token.Or,
                                   Token.Addition, Token.Subtraction, Token.Multiplication, Token.Division
                        If notLastToken Then FixSpaces(textEdit, line, t, tokens(i + 1), 1)

                    Case Token.LeftBracket, Token.LeftCurlyBracket, Token.LeftParens
                        If notLastToken Then FixSpaces(textEdit, line, t, tokens(i + 1), 0)

                    Case Token.RightBracket, Token.RightCurlyBracket, Token.RightParens
                        If notLastToken Then
                            Select Case tokens(i + 1).Token
                                Case Token.Comma, Token.LeftBracket, Token.LeftCurlyBracket, Token.LeftParens, Token.RightBracket, Token.RightCurlyBracket, Token.RightParens
                                    FixSpaces(textEdit, line, t, tokens(i + 1), 0)
                                Case Else
                                    FixSpaces(textEdit, line, t, tokens(i + 1), 1)
                            End Select
                        End If

                    Case Token.Comma
                        If notLastToken Then FixSpaces(textEdit, line, t, tokens(i + 1), 1)

                    Case Token.Identifier, Token.StringLiteral, Token.NumericLiteral
                        If notLastToken Then
                            Select Case tokens(i + 1).Token
                                Case Token.Dot, Token.Lookup, Token.Comma,
                                        Token.LeftBracket, Token.LeftCurlyBracket, Token.LeftParens,
                                        Token.RightBracket, Token.RightCurlyBracket, Token.RightParens
                                    FixSpaces(textEdit, line, t, tokens(i + 1), 0)
                                Case Else
                                    FixSpaces(textEdit, line, t, tokens(i + 1), 1)
                            End Select
                        End If

                    Case Else
                        If t.TokenType = TokenType.Keyword Then
                            If notLastToken Then FixSpaces(textEdit, line, t, tokens(i + 1), 1)
                        End If
                End Select
            Next
        End Sub

        Private Sub FixSpaces(
                       textEdit As ITextEdit,
                       line As ITextSnapshotLine,
                       token1 As TokenInfo,
                       token2 As TokenInfo,
                       requiredSpaces As Integer
                )

            If token2.Token = Token.Comment Then Return

            Dim spaces = token2.Column - token1.EndColumn
            If spaces < 0 OrElse spaces = requiredSpaces Then Return

            Dim n = spaces - requiredSpaces
            If n > 0 Then
                textEdit.Replace(line.Start + token2.Column - n, n, "")
            Else
                textEdit.Insert(line.Start + token2.Column, Space(n * -1))
            End If
        End Sub

        Private Sub AdjustIndentation(textEdit As ITextEdit, line As ITextSnapshotLine, indentationLevel As Integer, firstCharPos As Integer)
            If indentationLevel < 0 Then indentationLevel = 0
            If firstCharPos <> indentationLevel * 3 Then
                If firstCharPos = 19 Then Stop
                textEdit.Replace(
                    line.Start, firstCharPos,
                    New String(" "c, indentationLevel * 3)
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
    End Module
End Namespace
