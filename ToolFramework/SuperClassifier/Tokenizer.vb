Imports System.Collections.Generic
Imports Microsoft.Nautilus.Text
Imports System.Runtime.InteropServices

Namespace SuperClassifier
    Friend NotInheritable Class Tokenizer
        Friend Const NumberLiteralClassification As String = "NumberLiteral"
        Friend Const IdentifierClassification As String = "Identifier"
        Friend Const KeywordClassification As String = "Keyword"
        Friend Const UnknownClassification As String = "Unknown"
        Friend Const EOSClassification As String = "EndOfStream"

        Private languageData As LanguageData
        Private ReadOnly textBuffer As ITextBuffer
        Private tokenTree As New SplayTokenTree

        Friend Sub New(langData As LanguageData, textBuffer As ITextBuffer)
            languageData = langData
            Me.textBuffer = textBuffer
        End Sub

        Private Function GetReliableStart(start As Integer) As Integer
            Return If(tokenTree.GetPrevTokenForPosition(start)?.TokenStart, 0)
        End Function

        Private Function Edit(editStart As Integer, insertCount As Integer) As Span
            Return tokenTree.UpdateKeys(editStart, insertCount)
        End Function

        Private Function SkipWhiteSpaces(start As Integer) As Integer
            Dim length1 As Integer = textBuffer.CurrentSnapshot.Length
            While start < length1 AndAlso Char.IsWhiteSpace(textBuffer.CurrentSnapshot(start))
                start += 1
            End While

            Return start
        End Function

        Private Shared Function IsDigit(c As Char) As Boolean
            If "0"c <= c Then Return c <= "9"c
            Return False
        End Function

        Private Shared Function IsHexDigit(c As Char) As Boolean
            If Not IsDigit(c) AndAlso ("A"c > c OrElse c > "F"c) Then
                If "a"c <= c Then Return c <= "f"c
                Return False
            End If

            Return True
        End Function

        Private Function ScanNumber(start As Integer) As Token
            Dim num As Integer = start
            Dim c As Char = textBuffer.CurrentSnapshot(start)
            If Not IsDigit(c) Then Return Nothing

            Dim length1 As Integer = textBuffer.CurrentSnapshot.Length
            start += 1
            Dim c2 As Char

            If c = "0"c AndAlso start < length1 Then
                c2 = textBuffer.CurrentSnapshot(start)
                If c2 = "x"c OrElse c2 = "X"c Then
                    If (start + 1 < length1 AndAlso Not IsHexDigit(textBuffer.CurrentSnapshot(start + 1))) OrElse start + 1 >= length1 Then
                        Return New Token(num, start - num, textBuffer.CurrentSnapshot.GetText(num, start - num), "NumberLiteral")
                    End If

                    start += 1
                    While start < length1
                        c2 = textBuffer.CurrentSnapshot(start)
                        If Not IsHexDigit(c2) Then Exit While
                        start += 1
                    End While

                    start = ScanNumberSuffix(start)
                    Return New Token(num, start - num, textBuffer.CurrentSnapshot.GetText(num, start - num), "NumberLiteral")
                End If

            End If

            Dim flag As Boolean = False
            Dim flag2 As Boolean = False

            While start < length1
                c2 = textBuffer.CurrentSnapshot(start)
                If IsDigit(c2) Then Continue While

                Select Case c2
                    Case "."c
                        If flag Then Exit While
                        flag = True

                    Case "E"c, "e"c
                        If flag2 Then Exit While
                        flag2 = True
                        flag = True

                    Case "+"c, "-"c
                        Dim c3 As Char = textBuffer.CurrentSnapshot(start - 1)
                        If c3 <> "e"c AndAlso c3 <> "E"c Then Exit While

                    Case Else
                        Exit While
                End Select

                start += 1
            End While

            c2 = textBuffer.CurrentSnapshot(start - 1)
            Select Case c2
                Case "."c
                    start -= 1
                    Return New Token(num, start - num, textBuffer.CurrentSnapshot.GetText(num, start - num), "NumberLiteral")

                Case "+"c, "-"c
                    start -= 1
                    c2 = textBuffer.CurrentSnapshot(start - 1)

            End Select

            If c2 = "e"c OrElse c2 = "E"c Then start -= 1
            start = ScanNumberSuffix(start)
            Return New Token(num, start - num, textBuffer.CurrentSnapshot.GetText(num, start - num), "NumberLiteral")
        End Function

        Private Function ScanNumberSuffix(start As Integer) As Integer
            Dim length1 As Integer = textBuffer.CurrentSnapshot.Length
            If start >= length1 Then Return start

            Select Case textBuffer.CurrentSnapshot(start)
                Case "U"c, "u"c
                    start += 1
                    If start < length1 Then
                        Dim c2 As Char = textBuffer.CurrentSnapshot(start)
                        If c2 = "l"c OrElse c2 = "L"c Then start += 1
                    End If

                Case "L"c, "l"c
                    start += 1
                    If start < length1 Then
                        Dim c As Char = textBuffer.CurrentSnapshot(start)
                        If c = "u"c OrElse c = "U"c Then start += 1
                    End If

                Case "D"c, "F"c, "M"c, "d"c, "f"c, "m"c
                    start += 1

            End Select

            Return start

        End Function

        Private Function GetToken(start As Integer) As Token
            If languageData Is Nothing Then
                Dim lineFromPosition As ITextSnapshotLine = textBuffer.CurrentSnapshot.GetLineFromPosition(start)
                Dim text As String = textBuffer.CurrentSnapshot.GetText(start, lineFromPosition.EndIncludingLineBreak - start)
                Return New Token(start, lineFromPosition.EndIncludingLineBreak - start, text, "Unknown")
            End If

            start = SkipWhiteSpaces(start)
            If start >= textBuffer.CurrentSnapshot.Length Then
                Return New Token(start, 0, "<EOS>", "EndOfStream")
            End If

            Dim tokenStartingAt As Token
            For Each lexExpression As LexExpression In languageData.LexExpressions
                tokenStartingAt = lexExpression.GetTokenStartingAt(textBuffer.CurrentSnapshot, start)
                If tokenStartingAt IsNot Nothing Then
                    Return tokenStartingAt
                End If
            Next

            tokenStartingAt = languageData.IdentifierLexExpression.GetTokenStartingAt(textBuffer.CurrentSnapshot, start)
            If tokenStartingAt IsNot Nothing Then
                Return tokenStartingAt
            End If

            tokenStartingAt = ScanNumber(start)
            If tokenStartingAt IsNot Nothing Then
                Return tokenStartingAt
            End If

            Return New Token(start, 1, textBuffer.CurrentSnapshot(start).ToString(), "Unknown")
        End Function

        Private Function AddTokensToTree(tokenList As List(Of Token)) As Span
            If tokenList.Count = 0 Then
                Return New Span(0, 0)
            End If

            Dim tokenStart1 As Integer = tokenList(0).TokenStart
            Dim tokenEnd1 As Integer = tokenList(tokenList.Count - 1).TokenEnd
            Dim result As Span = tokenTree.RemoveRange(tokenStart1, tokenEnd1)
            Dim num As Integer = -1

            For Each token1 As Token In tokenList
                Dim lineNumberFromPosition As Integer = textBuffer.CurrentSnapshot.GetLineNumberFromPosition(token1.TokenStart)
                If num <> lineNumberFromPosition Then
                    num = lineNumberFromPosition
                    tokenTree.Insert(token1)
                End If
            Next

            If tokenList.Count > 0 Then
                tokenTree.Insert(tokenList(tokenList.Count - 1))
            End If

            Return result
        End Function

        Private Function TokenInTree(token As Token) As Boolean
            Dim prevToken As Token = tokenTree.GetPrevTokenForPosition(token.TokenStart)
            Return prevToken IsNot Nothing AndAlso
                        prevToken.TokenStart = token.TokenStart AndAlso
                        prevToken.TokenLength = token.TokenLength AndAlso
                        prevToken.TokenString = token.TokenString
        End Function

        Private Function GetTokensCovering(start As Integer, length1 As Integer, <Out> ByRef invalidatedSpan As Span) As List(Of Token)
            Dim list1 As New List(Of Token)
            Dim num As Integer = GetReliableStart(start)
            Dim length2 As Integer = textBuffer.CurrentSnapshot.Length
            Dim num2 As Integer = start + length1

            While num < length2
                Dim token1 As Token = GetToken(num)
                list1.Add(token1)
                num = token1.TokenEnd
                If num > num2 AndAlso (Not tokenTree.IsTokenEndAfter(num) OrElse TokenInTree(token1)) Then
                    Exit While
                End If
            End While

            invalidatedSpan = AddTokensToTree(list1)
            Return list1
        End Function

        Friend Function GetTokensCovering(start As Integer, length1 As Integer) As List(Of Token)
            Dim invalidatedSpan As Span
            Return GetTokensCovering(start, length1, invalidatedSpan)
        End Function

        Friend Function GetInvalidatedSpan(textChangeCollection As INormalizedTextChangeCollection) As TextSpan
            Dim num As Integer = Integer.MaxValue
            Dim num2 As Integer = Integer.MinValue

            For Each item As ITextChange In textChangeCollection
                Dim span1 As Span = Edit(item.Position, item.Delta)
                Dim invalidatedSpan As Span
                Dim tokensCovering As List(Of Token) = GetTokensCovering(item.Position, item.NewEnd - item.Position, invalidatedSpan)
                Dim span2 As Span = invalidatedSpan

                If tokensCovering.Count = 0 Then span2 = span1
                num = Math.Min(num, span2.Start)
                num2 = Math.Max(num2, span2.End)
            Next

            If num2 > textBuffer.CurrentSnapshot.Length Then
                num2 = textBuffer.CurrentSnapshot.Length
            End If

            Return New TextSpan(textBuffer.CurrentSnapshot, New Span(num, num2 - num), SpanTrackingMode.EdgeExclusive)
        End Function
    End Class
End Namespace
