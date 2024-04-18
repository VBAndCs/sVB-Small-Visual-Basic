Imports System.Globalization
Imports System.Text
Imports System.Runtime.InteropServices

Namespace Evaluator
    Friend Class LineScanner

        Private Shared _currentIndex As Integer
        Private Shared _lineLength As Integer
        Private Shared _lineText As String
        Private Shared _decimalSeparator As Char = "."c


        Public Shared Function GetTokenEnumerator(
                    lineText As String
                ) As TokenEnumerator

            Dim e As New TokenEnumerator(GetTokens(lineText, 0))
            e.LineNumber = 0
            Return e
        End Function

        Public Shared Function GetTokens(
                       lineText As String,
                       ByRef lineNumber As Integer
                  ) As List(Of Token)

            Dim tokens As New List(Of Token)()
            If lineText = "" Then Return tokens

            _lineText = lineText
            _lineLength = lineText.Length
            _currentIndex = 0

            Dim token As New Token
            Do
                If ScanNextToken(token) Then
                    token.Line = lineNumber
                    tokens.Add(token)
                Else
                    Exit Do
                End If
            Loop

            Return tokens
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

                Case "%"c
                    token.Type = TokenType.Mod
                    token.Text = "%"

                Case "*"c
                    token.Type = TokenType.Multiplication
                    token.Text = "*"

                Case "^"c
                    token.Type = TokenType.Power
                    token.Text = "^"

                Case "("c
                    token.Type = TokenType.LeftParens
                    token.Text = "("

                Case ")"c
                    token.Type = TokenType.RightParens
                    token.Text = ")"

                Case ","c
                    token.Type = TokenType.Comma
                    token.Text = ","

                Case Else
                    Dim nextChar2 As Char = GetNextChar()
                    _currentIndex -= 1

                    If Char.IsLetter(currentChar) OrElse currentChar = "_"c Then
                        _currentIndex -= 1
                        Dim word = ReadKeywordOrIdentifier()
                        token.Type = If(word = "mod", TokenType.Mod, TokenType.Identifier)
                        token.Text = word

                    ElseIf Char.IsDigit(currentChar) OrElse currentChar = "-"c OrElse currentChar = _decimalSeparator AndAlso Char.IsDigit(nextChar2) Then
                        _currentIndex -= 1
                        Dim literal As String = ReadNumericLiteral()
                        token.Type = TokenType.NumericLiteral
                        token.Text = literal

                    Else
                        _currentIndex -= 1
                        Dim illegal As String = ReadUntilNextSpace()
                        token.Type = TokenType.Illegal
                        token.Text = illegal
                    End If
            End Select

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

    End Class
End Namespace
