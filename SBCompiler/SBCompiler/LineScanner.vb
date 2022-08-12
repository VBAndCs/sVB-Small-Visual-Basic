Imports System
Imports System.Collections.Generic
Imports System.Globalization
Imports System.Text
Imports System.Runtime.InteropServices

Namespace Microsoft.SmallBasic
    Public Class LineScanner

        Private Shared _currentIndex As Integer
        Private Shared _lineLength As Integer
        Private Shared _lineText As String
        Private Shared _decimalSeparator As Char = "."c

        Public Shared Function GetFirstToken(lineText As String, lineNumber As Integer) As TokenInfo
            Dim tokens = GetTokens(lineText, lineNumber, Nothing, True)
            If tokens.Count = 0 Then Return TokenInfo.Illegal
            Return tokens(0)
        End Function

        Public Shared Function GetTokenEnumerator(
                    lineText As String,
                    ByRef lineNumber As Integer,
                    Optional lines As List(Of String) = Nothing
                ) As TokenEnumerator

            Dim e As New TokenEnumerator(GetTokens(lineText, lineNumber, lines))
            e.LineNumber = lineNumber
            Return e
        End Function

        Public Shared Function GetTokens(
                       lineText As String,
                       ByRef lineNumber As Integer,
                       Optional lines As List(Of String) = Nothing,
                       Optional firstTokenOnly As Boolean = False
                  ) As List(Of TokenInfo)

            Dim tokens As New List(Of TokenInfo)()
            If lineText = "" Then Return tokens

            _lineText = lineText
            _lineLength = lineText.Length
            _currentIndex = 0

            Dim tokenInfo As New TokenInfo
            Dim subLine = 0
            Do
                If ScanNextToken(tokenInfo) Then
                    tokenInfo.Line = lineNumber
                    tokenInfo.subLine = subLine
                    tokens.Add(tokenInfo)
                    If firstTokenOnly Then Exit Do

                ElseIf lines IsNot Nothing AndAlso tokens.Count > 0 AndAlso
                            lineNumber < lines.Count - 1 Then

                    If IsLineContinuity(tokens,
                              If(lineNumber < lines.Count - 1, lines(lineNumber + 1), "")) Then
                        ' Scan the next line as a part of this line

                        _lineText = lines(lineNumber + 1)
                        If _lineText = "" Then Exit Do

                        _lineLength = _lineText.Length
                        _currentIndex = 0
                        lineNumber += 1
                        subLine += 1
                    Else
                        Exit Do
                    End If

                Else
                    Exit Do
                End If
            Loop
            Return tokens
        End Function

        Public Shared Function IsLineContinuity(tokens As List(Of TokenInfo), nextLine As String) As Boolean
            Dim last = tokens.Count - 1
            If last = -1 Then Return False

            Dim comment = -1

            If tokens(last).TokenType = TokenType.Comment Then
                comment = last
                last -= 1
                If last = -1 Then Return False
            End If

            Dim nextLineFirstToken = GetFirstToken(nextLine, 0).NormalizedText

            Select Case tokens(last).NormalizedText
                Case "_"
                    If last = 0 OrElse tokens(last - 1).Text = "." OrElse nextLineFirstToken = "." Then
                        Return False
                    Else
                        tokens.RemoveAt(last) ' remove the Hyphen
                        comment -= 1
                    End If

                Case "-"
                    If tokens(0).Token = Token.ContinueLoop OrElse tokens(0).Token = Token.ExitLoop Then
                        Return False
                    End If
                Case ",", "{", "(", "[", "+", "*", "\", "/", "=", "or", "and"

                Case Else
                    Select Case nextLineFirstToken
                        Case ")", "]", "}", "+", "-", "*", "\", "/", "or", "and"

                        Case Else
                            Return False
                    End Select
            End Select

            If comment > -1 Then tokens.RemoveAt(comment)
            Return True
        End Function

        Private Shared Function ScanNextToken(<Out> ByRef tokenInfo As TokenInfo) As Boolean
            EatSpaces()
            tokenInfo = Nothing
            tokenInfo.Column = _currentIndex
            Dim nextChar As Char = GetNextChar()

            If nextChar = VisualBasic.Strings.ChrW(0) Then
                Return False
            End If

            Select Case nextChar
                Case "+"c
                    tokenInfo.Token = Token.Addition
                    tokenInfo.Text = "+"

                Case "-"c
                    tokenInfo.Token = Token.Subtraction
                    tokenInfo.Text = "-"

                Case "/"c
                    tokenInfo.Token = Token.Division
                    tokenInfo.Text = "/"

                Case "*"c
                    tokenInfo.Token = Token.Multiplication
                    tokenInfo.Text = "*"

                Case "("c
                    tokenInfo.Token = Token.LeftParens
                    tokenInfo.Text = "("

                Case ")"c
                    tokenInfo.Token = Token.RightParens
                    tokenInfo.Text = ")"

                Case "["c
                    tokenInfo.Token = Token.LeftBracket
                    tokenInfo.Text = "["

                Case "]"c
                    tokenInfo.Token = Token.RightBracket
                    tokenInfo.Text = "]"

                Case "{"c
                    tokenInfo.Token = Token.LeftCurlyBracket
                    tokenInfo.Text = "{"

                Case "}"c
                    tokenInfo.Token = Token.RightCurlyBracket
                    tokenInfo.Text = "}"

                Case ":"c
                    tokenInfo.Token = Token.Colon
                    tokenInfo.Text = ":"

                Case ","c
                    tokenInfo.Token = Token.Comma
                    tokenInfo.Text = ","

                Case "<"c
                    Select Case GetNextChar()
                        Case "="c
                            tokenInfo.Token = Token.LessThanEqualTo
                            tokenInfo.Text = "<="
                        Case ">"c
                            tokenInfo.Token = Token.NotEqualTo
                            tokenInfo.Text = "<>"
                        Case Else
                            _currentIndex -= 1
                            tokenInfo.Token = Token.LessThan
                            tokenInfo.Text = "<"
                    End Select

                Case ">"c
                    nextChar = GetNextChar()
                    If nextChar = "="c Then
                        tokenInfo.Token = Token.GreaterThanEqualTo
                        tokenInfo.Text = ">="
                    Else
                        _currentIndex -= 1
                        tokenInfo.Token = Token.GreaterThan
                        tokenInfo.Text = ">"
                    End If

                Case "="c
                    tokenInfo.Token = Token.Equals
                    tokenInfo.Text = "="

                Case "'"c
                    _currentIndex -= 1
                    tokenInfo.Token = Token.Comment
                    tokenInfo.Text = ReadComment()

                Case """"c
                    _currentIndex -= 1
                    tokenInfo.Token = Token.StringLiteral
                    tokenInfo.Text = ReadStringLiteral()

                Case Else
                    Dim nextChar2 As Char = GetNextChar()
                    _currentIndex -= 1

                    If Char.IsLetter(nextChar) OrElse nextChar = "_"c Then
                        _currentIndex -= 1
                        Dim text As String = ReadKeywordOrIdentifier()
                        tokenInfo.Token = GetToken(text)
                        tokenInfo.Text = text

                    ElseIf Char.IsDigit(nextChar) OrElse nextChar = "-"c OrElse nextChar = _decimalSeparator AndAlso Char.IsDigit(nextChar2) Then
                        _currentIndex -= 1
                        Dim text2 As String = ReadNumericLiteral()
                        tokenInfo.Token = Token.NumericLiteral
                        tokenInfo.Text = text2

                    ElseIf nextChar = "."c Then
                        tokenInfo.Token = Token.Dot
                        tokenInfo.Text = "."

                    ElseIf nextChar = "!"c Then
                        tokenInfo.Token = Token.Lookup
                        tokenInfo.Text = "!"

                    Else
                        _currentIndex -= 1
                        Dim text3 As String = ReadUntilNextSpace()
                        tokenInfo.Token = Token.Illegal
                        tokenInfo.TokenType = TokenType.Illegal
                        tokenInfo.Text = text3
                    End If

                    Exit Select
            End Select

            tokenInfo.TokenType = GetTokenType(tokenInfo.Token)
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
            Dim stringBuilder As StringBuilder = New StringBuilder()
            Dim nextChar As Char = GetNextChar()

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

        Private Shared Function ReadStringLiteral() As String
            Dim stringBuilder As StringBuilder = New StringBuilder()
            Dim nextChar As Char = GetNextChar()
            stringBuilder.Append(nextChar)
            nextChar = GetNextChar()
            Dim flag = False

            While AscW(nextChar) <> 0
                stringBuilder.Append(nextChar)

                If nextChar = """"c Then
                    nextChar = GetNextChar()

                    If nextChar <> """"c Then
                        flag = True
                        Exit While
                    End If

                    stringBuilder.Append(nextChar)
                End If

                nextChar = GetNextChar()
            End While

            If flag Then
                _currentIndex -= 1
                Return stringBuilder.ToString()
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

        Private Shared Function IsWhiteSpace(ByVal c As Char) As Boolean
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

        Public Shared Function GetToken(tokenText As String) As Token
            Select Case tokenText.ToLower(CultureInfo.InvariantCulture)
                Case "and"
                    Return Token.And
                Case "else"
                    Return Token.Else
                Case "elseif"
                    Return Token.ElseIf
                Case "endfor"
                    Return Token.EndFor
                Case "next"
                    Return Token.Next
                Case "endif"
                    Return Token.EndIf
                Case "endsub"
                    Return Token.EndSub
                Case "endfunction"
                    Return Token.EndFunction
                Case "return"
                    Return Token.Return
                Case "exitloop"
                    Return Token.ExitLoop
                Case "continueloop"
                    Return Token.ContinueLoop
                Case "endwhile"
                    Return Token.EndWhile
                Case "wend"
                    Return Token.Wend
                Case "for"
                    Return Token.For
                Case "goto"
                    Return Token.Goto
                Case "if"
                    Return Token.If
                Case "or"
                    Return Token.Or
                Case "step"
                    Return Token.Step
                Case "sub"
                    Return Token.Sub
                Case "function"
                    Return Token.Function
                Case "then"
                    Return Token.Then
                Case "to"
                    Return Token.To
                Case "while"
                    Return Token.While
                Case "true"
                    Return Token.True
                Case "false"
                    Return Token.False
                Case Else
                    Return Token.Identifier
            End Select
        End Function

        Public Shared Function GetTokenType(token As Token) As TokenType
            Select Case token
                Case Token.Illegal
                    Return TokenType.Illegal

                Case Token.Comment
                    Return TokenType.Comment

                Case Token.StringLiteral
                    Return TokenType.StringLiteral

                Case Token.NumericLiteral
                    Return TokenType.NumericLiteral

                Case Token.Identifier
                    Return TokenType.Identifier

                Case Token.Else, Token.ElseIf, Token.EndFor, Token.Next, Token.EndIf, Token.EndSub, Token.EndFunction, Token.EndWhile, Token.Wend, Token.For, Token.Goto, Token.Return, Token.ExitLoop, Token.ContinueLoop, Token.If, Token.Step, Token.Sub, Token.Function, Token.Then, Token.To, Token.While, Token.True, Token.False
                    Return TokenType.Keyword

                Case Token.And, Token.Equals, Token.Or, Token.Dot, Token.Lookup, Token.Addition, Token.Subtraction, Token.Division, Token.Multiplication, Token.LeftParens, Token.RightParens, Token.LessThan, Token.LessThanEqualTo, Token.GreaterThan, Token.GreaterThanEqualTo, Token.NotEqualTo, Token.Comma, Token.Colon
                    Return TokenType.Operator

                Case Else
                    Return TokenType.Illegal
            End Select
        End Function

    End Class
End Namespace
