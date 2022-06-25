Imports System.Collections.Generic
Imports System.Text
Imports Microsoft.Nautilus.Text

Namespace SuperClassifier
    Friend NotInheritable Class IdentifierLexExpression
        Inherits LexExpression

        Friend ReadOnly prefixChars As New Dictionary(Of Char, Char)
        Friend ReadOnly bodyChars As New Dictionary(Of Char, Char)
        Friend ReadOnly keywords As New Dictionary(Of String, String)
        Friend ReadOnly IsCaseSensitive As Boolean

        Friend Sub New(prefixChars As String, bodyChars As String, isCaseSensitive As Boolean)
            MyBase.New(LexExpressionKind.Identifier, "Identifier")
            If prefixChars IsNot Nothing Then
                For Each c As Char In prefixChars
                    Me.prefixChars(c) = c
                Next
            End If

            If bodyChars IsNot Nothing Then
                For Each c As Char In bodyChars
                    Me.bodyChars(c) = c
                Next
            End If
            Me.IsCaseSensitive = isCaseSensitive
        End Sub

        Friend Sub AddKeywords(keywords As IList(Of String))
            For Each keyword As String In keywords
                Dim text As String = keyword
                If Not IsCaseSensitive Then
                    text = keyword.ToUpperInvariant()
                End If
                Me.keywords(text) = text
            Next
        End Sub

        Friend Overrides Function GetTokenStartingAt(textBuffer As ITextSnapshot, start As Integer) As Token
            Dim c As Char = textBuffer(start)
            If Not prefixChars.ContainsKey(c) AndAlso Not bodyChars.ContainsKey(c) AndAlso Not Char.IsLetter(c) Then
                Return Nothing
            End If

            Dim num As Integer = start
            start += 1
            Dim stringBuilder1 As New StringBuilder
            stringBuilder1.Append(c)
            Dim length1 As Integer = textBuffer.Length

            While start < length1
                c = textBuffer(start)
                If Not bodyChars.ContainsKey(c) AndAlso Not Char.IsLetterOrDigit(c) Then
                    Exit While
                End If
                stringBuilder1.Append(c)
                start += 1
            End While

            Dim text As String = stringBuilder1.ToString()
            Dim key As String = text
            If Not IsCaseSensitive Then
                key = text.ToUpperInvariant()
            End If

            If keywords.ContainsKey(key) Then
                Return New Token(num, start - num, stringBuilder1.ToString(), "Keyword")
            End If

            Return New Token(num, start - num, stringBuilder1.ToString(), "Identifier")
        End Function
    End Class
End Namespace
