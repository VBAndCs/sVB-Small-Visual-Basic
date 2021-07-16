Imports System
Imports System.Collections.Generic
Imports System.Globalization
Imports System.Text
Imports System.Runtime.InteropServices

Namespace Microsoft.SmallBasic
    Public Class LineScanner
        Private _currentIndex As Integer
        Private _lineLength As Integer
        Private _lineText As String
        Private _decimalSeparator As Char = "."c

        Public Function GetTokenList(lineText As String, lineNumber As Integer) As TokenEnumerator
            Dim tokenEnumerator As New TokenEnumerator(GetTokens(lineText, lineNumber))
            tokenEnumerator.LineNumber = lineNumber
            Return tokenEnumerator
        End Function

        Public Function GetTokens(lineText As String, lineNumber As Integer) As List(Of TokenInfo)
            If Equals(lineText, Nothing) Then
                Throw New ArgumentNullException("lineText")
            End If

            _lineText = lineText
            _lineLength = _lineText.Length
            _currentIndex = 0
            Dim list As New List(Of TokenInfo)()
            Dim tokenInfo As New TokenInfo

            While ScanNextToken(tokenInfo)
                tokenInfo.Line = lineNumber
                list.Add(tokenInfo)
            End While

            Return list
        End Function

        Private Function ScanNextToken(<Out> ByRef tokenInfo As TokenInfo) As Boolean
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
                        tokenInfo.Token = MatchToken(text)
                        tokenInfo.Text = text
                    ElseIf Char.IsDigit(nextChar) OrElse nextChar = "-"c OrElse nextChar = _decimalSeparator AndAlso Char.IsDigit(nextChar2) Then
                        _currentIndex -= 1
                        Dim text2 As String = ReadNumericLiteral()
                        tokenInfo.Token = Token.NumericLiteral
                        tokenInfo.Text = text2
                    ElseIf nextChar = "."c Then
                        tokenInfo.Token = Token.Dot
                        tokenInfo.Text = "."
                    Else
                        _currentIndex -= 1
                        Dim text3 As String = ReadUntilNextSpace()
                        tokenInfo = TokenInfo.Illegal
                        tokenInfo.Text = text3
                    End If

                    Exit Select
            End Select

            tokenInfo.TokenType = GetTokenType(tokenInfo.Token)
            Return True
        End Function

        Private Function GetNextChar() As Char
            If _currentIndex < _lineLength Then
                Return _lineText(Math.Min(Threading.Interlocked.Increment(_currentIndex), _currentIndex - 1))
            End If

            _currentIndex += 1
            Return VisualBasic.Strings.ChrW(0)
        End Function

        Private Function ReadUntilNextSpace() As String
            Dim stringBuilder As StringBuilder = New StringBuilder()
            Dim nextChar As Char = GetNextChar()

            While Not IsWhiteSpace(nextChar) AndAlso AscW(nextChar) <> 0
                stringBuilder.Append(nextChar)
                nextChar = GetNextChar()
            End While

            _currentIndex -= 1
            Return stringBuilder.ToString()
        End Function

        Private Function ReadKeywordOrIdentifier() As String
            Dim stringBuilder As StringBuilder = New StringBuilder()
            Dim nextChar As Char = GetNextChar()

            While Char.IsLetterOrDigit(nextChar) OrElse nextChar = "_"c
                stringBuilder.Append(nextChar)
                nextChar = GetNextChar()
            End While

            _currentIndex -= 1
            Return stringBuilder.ToString()
        End Function

        Private Function ReadNumericLiteral() As String
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

        Private Function ReadStringLiteral() As String
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

        Private Function ReadComment() As String
            Dim stringBuilder As StringBuilder = New StringBuilder()
            Dim nextChar As Char = GetNextChar()
            stringBuilder.Append(nextChar)
            stringBuilder.Append(_lineText.Substring(_currentIndex))
            _currentIndex = _lineLength
            Return stringBuilder.ToString()
        End Function

        Private Sub EatSpaces()
            While IsWhiteSpace(GetNextChar())
            End While

            _currentIndex -= 1
        End Sub

        Private Function IsWhiteSpace(ByVal c As Char) As Boolean
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

        Private Function MatchToken(ByVal tokenText As String) As Token
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
                Case "then"
                    Return Token.Then
                Case "to"
                    Return Token.To
                Case "while"
                    Return Token.While
                Case Else
                    Return Token.Identifier
            End Select
        End Function

        Private Function GetTokenType(ByVal token As Token) As TokenType
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
                Case Token.Else, Token.ElseIf, Token.EndFor, Token.Next, Token.EndIf, Token.EndSub, Token.EndWhile, Token.Wend, Token.For, Token.Goto, Token.If, Token.Step, Token.Sub, Token.Then, Token.To, Token.While
                    Return TokenType.Keyword
                Case Token.And, Token.Equals, Token.Or, Token.Dot, Token.Addition, Token.Subtraction, Token.Division, Token.Multiplication, Token.LeftParens, Token.RightParens, Token.LessThan, Token.LessThanEqualTo, Token.GreaterThan, Token.GreaterThanEqualTo, Token.NotEqualTo, Token.Comma, Token.Colon
                    Return TokenType.Operator
                Case Else
                    Return TokenType.Illegal
            End Select
        End Function

    End Class
End Namespace
