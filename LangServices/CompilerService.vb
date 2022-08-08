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

        Public Sub FormatDocument(textBuffer As ITextBuffer, Optional lineNumber As Integer = -1)
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

                Dim nextLineOffset = 0
                Dim resetNextLineOffset = False

                For i = start To [end]
                    Dim line = currentSnapshot.Lines(i)
                    Dim nextPos = line.GetPositionOfNextNonWhiteSpaceCharacter(0)
                    Dim tokens = LineScanner.GetTokens(line.GetText(), line.LineNumber)

                    If tokens.Count = 0 Then
                        AdjustIndentation(textEdit, line, indentationLevel + nextLineOffset, nextPos)
                        nextLineOffset = 0
                        Continue For
                    End If

                    Select Case tokens(0).Token
                        Case Token.EndIf, Token.Next, Token.EndFor, Token.Wend, Token.EndWhile
                            indentationLevel -= 1
                            AdjustIndentation(textEdit, line, indentationLevel, nextPos)

                        Case Token.EndSub, Token.EndFunction
                            indentationLevel = 0
                            AdjustIndentation(textEdit, line, indentationLevel, nextPos)

                        Case Token.If, Token.For, Token.While
                            AdjustIndentation(textEdit, line, indentationLevel, nextPos)
                            indentationLevel += 1

                        Case Token.Sub, Token.Function
                            AdjustIndentation(textEdit, line, 0, nextPos)
                            indentationLevel = 1

                        Case Token.Else, Token.ElseIf
                            indentationLevel -= 1
                            AdjustIndentation(textEdit, line, indentationLevel, nextPos)
                            indentationLevel += 1

                        Case Token.RightParens, Token.RightBracket, Token.RightCurlyBracket
                            nextLineOffset = Math.Max(0, nextLineOffset - 1)
                            AdjustIndentation(textEdit, line, indentationLevel + nextLineOffset, nextPos)

                        Case Token.Addition, Token.Subtraction, Token.Multiplication, Token.Division, Token.Or, Token.And
                            If nextLineOffset = 0 Then nextLineOffset = 1
                            AdjustIndentation(textEdit, line, indentationLevel + nextLineOffset, nextPos)

                        Case Else
                            If resetNextLineOffset Then nextLineOffset = 0
                            AdjustIndentation(textEdit, line, indentationLevel + nextLineOffset, nextPos)
                    End Select

                    resetNextLineOffset = False

                    Dim last = tokens.Count - 1
                    If tokens(last).TokenType = TokenType.Comment Then
                        last -= 1
                    End If

                    If last = -1 Then
                        nextLineOffset = 0
                        Continue For
                    End If

                    Select Case tokens(last).NormalizedText
                        Case "_", ",", "+", "-", "*", "/", "and", "or"
                            If nextLineOffset = 0 Then nextLineOffset = 1
                        Case "(", "{", "[", "="
                            nextLineOffset += 1
                        Case ")", "}", "]"
                            nextLineOffset = Math.Max(0, nextLineOffset - 1)
                        Case Else
                            resetNextLineOffset = True
                    End Select
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


        Private Sub AdjustIndentation(textEdit As ITextEdit, line As ITextSnapshotLine, indentationLevel As Integer, nextPos As Integer)
            If indentationLevel < 0 Then indentationLevel = 0
            If nextPos <> indentationLevel * 3 Then
                textEdit.Replace(line.Start, nextPos, New String(" "c, indentationLevel * 3))
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
