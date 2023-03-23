Imports System.Globalization
Imports System.Text
Imports System.Runtime.InteropServices

Namespace Microsoft.SmallVisualBasic
    Public Class LineScanner

        Private Shared _currentIndex As Integer
        Private Shared _lineLength As Integer
        Private Shared _lineText As String
        Private Shared _decimalSeparator As Char = "."c

        Public Shared Function GetFirstToken(lineText As String, lineNumber As Integer) As Token
            Dim tokens = GetTokens(lineText, lineNumber, Nothing, 1)
            If tokens.Count = 0 Then Return Token.Illegal
            Return tokens(0)
        End Function

        Public Shared Function GetFirstTokens(lineText As String, lineNumber As Integer, maxTokenCount As Integer) As List(Of Token)
            Dim tokens = GetTokens(lineText, lineNumber, Nothing, maxTokenCount)
            Return tokens
        End Function

        Public Shared Function GetTokenEnumerator(
                    lineText As String,
                    ByRef lineNumber As Integer,
                    Optional lines As List(Of String) = Nothing,
                    Optional lineOffset As Integer = 0
                ) As TokenEnumerator

            Dim e As New TokenEnumerator(GetTokens(lineText, lineNumber, lines, False, lineOffset))
            e.LineNumber = lineNumber
            Return e
        End Function

        Public Shared Function GetTokens(
                       lineText As String,
                       ByRef lineNumber As Integer,
                       Optional lines As List(Of String) = Nothing,
                       Optional maxTokenCount As Integer = 0,
                       Optional lineOffset As Integer = 0
                  ) As List(Of Token)

            Dim tokens As New List(Of Token)()
            If lineText = "" Then Return tokens

            _lineText = lineText
            _lineLength = lineText.Length
            _currentIndex = 0

            Dim token As New Token
            Dim subLine = 0
            Do
                If ScanNextToken(token) Then
                    token.Line = lineNumber + lineOffset
                    token.subLine = subLine
                    tokens.Add(token)
                    If maxTokenCount > 0 AndAlso tokens.Count = maxTokenCount Then Exit Do

                ElseIf lines IsNot Nothing AndAlso tokens.Count > 0 AndAlso
                            lineNumber < lines.Count - 1 Then

                    If IsLineContinuity(tokens,
                              If(lineNumber < lines.Count - 1,
                                  lines(lineNumber + 1),
                                  ""
                              )
                         ) Then
                        ' Scan the next line as a part of this line
                        lineNumber += 1
                        _lineText = lines(lineNumber)
                        _lineLength = _lineText.Length
                        _currentIndex = 0
                        subLine += 1

                        If _lineText.Trim() = "" Then
                            token = Token.Illegal
                            token.Line = lineNumber + lineOffset
                            token.Column = _lineText.Length
                            token.subLine = subLine
                            tokens.Add(token)
                            Exit Do
                        End If

                    Else
                        Exit Do
                    End If

                Else
                    Exit Do
                End If
            Loop
            Return tokens
        End Function

        Public Shared SubLineComments As New List(Of Token)
        Friend Shared LineOffset As Integer

        Public Shared Function IsLineContinuity(
                             tokens As List(Of Token),
                             nextLine As String
                     ) As Boolean

            Dim last = tokens.Count - 1
            If last = -1 Then Return False

            Dim comment = -1
            Dim currentToken = tokens(last)
            Dim lastToken As Token

            If currentToken.ParseType = ParseType.Comment Then
                comment = last
                last -= 1
                If last = -1 Then Return False
            End If

            lastToken = tokens(last)

            ' comment lines are not allowed as sublines
            If currentToken.Line > lastToken.Line Then Return False

            Dim nextLineFirstToken = GetFirstToken(nextLine, 0)
            Dim nextLineFirstTokenText = nextLineFirstToken.LCaseText

            Select Case lastToken.LCaseText
                Case "_"
                    If last = 0 OrElse tokens(last - 1).Text = "." OrElse nextLineFirstTokenText = "." Then
                        Return False
                    Else
                        tokens.RemoveAt(last) ' remove the Hyphen
                        comment -= 1
                    End If

                Case "-"
                    If tokens(0).Type = TokenType.ContinueLoop OrElse tokens(0).Type = TokenType.ExitLoop Then
                        Return False
                    End If

                Case ",", "{", "(", "[", "+", "*", "\", "/", "=", "or", "and"
                    Select Case nextLineFirstToken.Type
                        Case TokenType.If, TokenType.Else, TokenType.ElseIf, TokenType.EndIf,
                                 TokenType.Goto, TokenType.ExitLoop, TokenType.ContinueLoop, TokenType.Return,
                                 TokenType.While, TokenType.EndWhile, TokenType.Wend,
                                 TokenType.For, TokenType.ForEach, TokenType.Next, TokenType.EndFor,
                                 TokenType.Sub, TokenType.EndSub, TokenType.Function, TokenType.EndFunction
                            Return False ' a start or end of a block can't be a line continuty
                    End Select

                Case Else
                    Select Case nextLineFirstTokenText
                        Case ")", "]", "}", "+", "-", "*", "\", "/", "or", "and"

                        Case Else
                            Return False
                    End Select
            End Select

            If comment > -1 Then
                SubLineComments.Add(tokens(comment))
                tokens.RemoveAt(comment)
            End If

            Return True
        End Function

        Private Shared Function ScanNextToken(<Out> ByRef token As Token) As Boolean
            EatSpaces()
            token = Nothing
            token.Column = _currentIndex
            Dim currentChar As Char = GetNextChar()

            If currentChar = ChrW(0) Then
                Return False
            End If

            Select Case currentChar
                Case "+"c
                    token.Type = TokenType.Addition
                    token.Text = "+"

                Case "-"c
                    token.Type = TokenType.Subtraction
                    token.Text = "-"

                Case "/"c
                    token.Type = TokenType.Division
                    token.Text = "/"

                Case "*"c
                    token.Type = TokenType.Multiplication
                    token.Text = "*"

                Case "("c
                    token.Type = TokenType.LeftParens
                    token.Text = "("

                Case ")"c
                    token.Type = TokenType.RightParens
                    token.Text = ")"

                Case "["c
                    token.Type = TokenType.LeftBracket
                    token.Text = "["

                Case "]"c
                    token.Type = TokenType.RightBracket
                    token.Text = "]"

                Case "{"c
                    token.Type = TokenType.LeftCurlyBracket
                    token.Text = "{"

                Case "}"c
                    token.Type = TokenType.RightCurlyBracket
                    token.Text = "}"

                Case ":"c
                    token.Type = TokenType.Colon
                    token.Text = ":"

                Case ","c
                    token.Type = TokenType.Comma
                    token.Text = ","

                Case "<"c
                    Select Case GetNextChar()
                        Case "="c
                            token.Type = TokenType.LessThanEqualTo
                            token.Text = "<="
                        Case ">"c
                            token.Type = TokenType.NotEqualTo
                            token.Text = "<>"
                        Case Else
                            _currentIndex -= 1
                            token.Type = TokenType.LessThan
                            token.Text = "<"
                    End Select

                Case ">"c
                    currentChar = GetNextChar()
                    If currentChar = "="c Then
                        token.Type = TokenType.GreaterThanEqualTo
                        token.Text = ">="
                    Else
                        _currentIndex -= 1
                        token.Type = TokenType.GreaterThan
                        token.Text = ">"
                    End If

                Case "="c
                    token.Type = TokenType.Equals
                    token.Text = "="

                Case "'"c
                    _currentIndex -= 1
                    token.Type = TokenType.Comment
                    token.Text = ReadComment()

                Case """"c
                    _currentIndex -= 1
                    token.Type = TokenType.StringLiteral
                    token.Text = ReadLiteral(currentChar)

                Case "#"c
                    _currentIndex -= 1
                    token.Type = TokenType.DateLiteral
                    token.Text = ReadLiteral(currentChar)

                Case Else
                    Dim nextChar2 As Char = GetNextChar()
                    _currentIndex -= 1

                    If Char.IsLetter(currentChar) OrElse currentChar = "_"c Then
                        _currentIndex -= 1
                        Dim text As String = ReadKeywordOrIdentifier()
                        token.Type = GetTokenType(text)
                        token.Text = text

                    ElseIf Char.IsDigit(currentChar) OrElse currentChar = "-"c OrElse currentChar = _decimalSeparator AndAlso Char.IsDigit(nextChar2) Then
                        _currentIndex -= 1
                        Dim text2 As String = ReadNumericLiteral()
                        token.Type = TokenType.NumericLiteral
                        token.Text = text2

                    ElseIf currentChar = "."c Then
                        token.Type = TokenType.Dot
                        token.Text = "."

                    ElseIf currentChar = "!"c Then
                        token.Type = TokenType.Lookup
                        token.Text = "!"

                    Else
                        _currentIndex -= 1
                        Dim text3 As String = ReadUntilNextSpace()
                        token.Type = TokenType.Illegal
                        token.ParseType = ParseType.Illegal
                        token.Text = text3
                    End If

                    Exit Select
            End Select

            token.ParseType = GetParseType(token.Type)
            Dim x = token.Text

            token.Hidden = x.StartsWith("_") AndAlso (
                                      x.StartsWith("__foreach__") OrElse
                                      x.StartsWith("__tmpArray__") OrElse
                                      x.StartsWith("_sVB_dynamic_Data_")
                                   )
            Return True
        End Function

        Private Shared Function GetNextChar() As Char
            If _currentIndex < _lineLength Then
                GetNextChar = _lineText(_currentIndex)
            Else
                GetNextChar = ChrW(0)
            End If

            _currentIndex += 1
        End Function

        Private Shared Function ReadUntilNextSpace() As String
            Dim stringBuilder As New StringBuilder()
            Dim nextChar = GetNextChar()

            While Not IsWhiteSpace(nextChar) AndAlso AscW(nextChar) <> 0
                stringBuilder.Append(nextChar)
                nextChar = GetNextChar()
            End While

            _currentIndex -= 1
            Return stringBuilder.ToString()
        End Function

        Private Shared Function ReadKeywordOrIdentifier() As String
            Dim stringBuilder As StringBuilder = New StringBuilder()
            Dim nextChar As Char = GetNextChar()

            While Char.IsLetterOrDigit(nextChar) OrElse nextChar = "_"c
                stringBuilder.Append(nextChar)
                nextChar = GetNextChar()
            End While

            _currentIndex -= 1
            Return stringBuilder.ToString()
        End Function

        Private Shared Function ReadNumericLiteral() As String
            Dim stringBuilder As StringBuilder = New StringBuilder()
            Dim nextChar As Char = GetNextChar()
            GetNextChar()
            _currentIndex -= 1
            Dim flag = True
            Dim flag2 = False

            While nextChar <> "-"c OrElse flag

                If nextChar = _decimalSeparator AndAlso Not flag2 Then
                    flag2 = True
                ElseIf Not Char.IsDigit(nextChar) Then
                    Exit While
                End If

                flag = False
                stringBuilder.Append(nextChar)
                nextChar = GetNextChar()
            End While

            _currentIndex -= 1
            Return stringBuilder.ToString()
        End Function

        Private Shared Function ReadLiteral(closingChar As Char) As String
            Dim stringBuilder As New StringBuilder()
            Dim currentChar = GetNextChar()
            stringBuilder.Append(currentChar)
            currentChar = GetNextChar()
            Dim flag = False

            While AscW(currentChar) <> 0
                stringBuilder.Append(currentChar)

                If currentChar = closingChar Then
                    currentChar = GetNextChar()

                    If currentChar <> closingChar Then
                        flag = True
                        Exit While
                    End If

                    stringBuilder.Append(currentChar)
                End If

                currentChar = GetNextChar()
            End While

            If flag Then
                _currentIndex -= 1
            End If

            Return stringBuilder.ToString()
        End Function

        Private Shared Function ReadComment() As String
            Dim stringBuilder As StringBuilder = New StringBuilder()
            Dim nextChar As Char = GetNextChar()
            stringBuilder.Append(nextChar)
            stringBuilder.Append(_lineText.Substring(_currentIndex))
            _currentIndex = _lineLength
            Return stringBuilder.ToString()
        End Function

        Private Shared Sub EatSpaces()
            While IsWhiteSpace(GetNextChar())
            End While

            _currentIndex -= 1
        End Sub

        Private Shared Function IsWhiteSpace(c As Char) As Boolean
            Select Case c
                Case VisualBasic.Strings.ChrW(9), VisualBasic.Strings.ChrW(11), VisualBasic.Strings.ChrW(12), VisualBasic.Strings.ChrW(26), " "c
                    Return True
                Case Else

                    If c >= VisualBasic.Strings.ChrW(128) Then
                        Return Char.GetUnicodeCategory(c) = UnicodeCategory.SpaceSeparator
                    End If

                    Return False
            End Select
        End Function

        Public Shared Function GetTokenType(tokenText As String) As TokenType
            Select Case tokenText.ToLower(CultureInfo.InvariantCulture)
                Case "and"
                    Return TokenType.And
                Case "else"
                    Return TokenType.Else
                Case "elseif"
                    Return TokenType.ElseIf
                Case "endfor"
                    Return TokenType.EndFor
                Case "next"
                    Return TokenType.Next
                Case "endif"
                    Return TokenType.EndIf
                Case "endsub"
                    Return TokenType.EndSub
                Case "endfunction"
                    Return TokenType.EndFunction
                Case "return"
                    Return TokenType.Return
                Case "exitloop"
                    Return TokenType.ExitLoop
                Case "continueloop"
                    Return TokenType.ContinueLoop
                Case "endwhile"
                    Return TokenType.EndWhile
                Case "wend"
                    Return TokenType.Wend
                Case "for"
                    Return TokenType.For
                Case "foreach"
                    Return TokenType.ForEach
                Case "goto"
                    Return TokenType.Goto
                Case "if"
                    Return TokenType.If
                Case "in"
                    Return TokenType.In
                Case "or"
                    Return TokenType.Or
                Case "step"
                    Return TokenType.Step
                Case "sub"
                    Return TokenType.Sub
                Case "function"
                    Return TokenType.Function
                Case "then"
                    Return TokenType.Then
                Case "to"
                    Return TokenType.To
                Case "while"
                    Return TokenType.While
                Case "true"
                    Return TokenType.True
                Case "false"
                    Return TokenType.False
                Case Else
                    Return TokenType.Identifier
            End Select
        End Function

        Public Shared Function GetParseType(token As TokenType) As ParseType
            Select Case token
                Case TokenType.Illegal
                    Return ParseType.Illegal

                Case TokenType.Comment
                    Return ParseType.Comment

                Case TokenType.StringLiteral
                    Return ParseType.StringLiteral

                Case TokenType.DateLiteral
                    Return ParseType.DateLiteral

                Case TokenType.NumericLiteral
                    Return ParseType.NumericLiteral

                Case TokenType.Identifier
                    Return ParseType.Identifier

                Case TokenType.Else, TokenType.ElseIf,
                         TokenType.EndFor, TokenType.Next,
                         TokenType.EndIf, TokenType.EndSub,
                         TokenType.EndFunction, TokenType.EndWhile,
                         TokenType.Wend, TokenType.For, TokenType.ForEach,
                         TokenType.Goto, TokenType.Return,
                         TokenType.ExitLoop, TokenType.ContinueLoop,
                         TokenType.If, TokenType.Step,
                         TokenType.Sub, TokenType.Function,
                         TokenType.Then, TokenType.To, TokenType.In, TokenType.While,
                         TokenType.True, TokenType.False
                    Return ParseType.Keyword

                Case TokenType.And, TokenType.Or,
                         TokenType.Equals, TokenType.NotEqualTo,
                         TokenType.LessThan, TokenType.LessThanEqualTo,
                         TokenType.GreaterThan, TokenType.GreaterThanEqualTo,
                         TokenType.Addition, TokenType.Subtraction,
                         TokenType.Division, TokenType.Multiplication,
                         TokenType.Dot, TokenType.Lookup, TokenType.Comma, TokenType.Colon,
                         TokenType.LeftParens, TokenType.RightParens,
                         TokenType.LeftBracket, TokenType.RightBracket,
                         TokenType.LeftCurlyBracket, TokenType.RightCurlyBracket
                    Return ParseType.Operator

                Case Else
                    Return ParseType.Illegal
            End Select
        End Function


    End Class
End Namespace
