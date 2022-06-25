Imports System.Collections.Generic
Imports System.Text
Imports Microsoft.Nautilus.Text

Namespace SuperClassifier
    Friend NotInheritable Class LiteralLexExpression
        Inherits LexExpression

        Friend ReadOnly startLiteralLookup As New Dictionary(Of Char, Dictionary(Of String, String))

        Friend Sub New(literalList As IList(Of String), classification As String)
            MyBase.New(LexExpressionKind.Literal, classification)
            For Each literal1 As String In literalList
                Dim key As Char = literal1(0)
                Dim _value As Collections.Generic.Dictionary(Of String, String) = Nothing
                If Not startLiteralLookup.TryGetValue(key, _value) Then
                    _value = New Dictionary(Of String, String)
                    startLiteralLookup.Add(key, _value)
                End If
                _value.Add(literal1, literal1)
            Next
        End Sub

        Friend Overrides Function GetTokenStartingAt(textBuffer As ITextSnapshot, start As Integer) As Token
            Dim c As Char = textBuffer(start)
            Dim _value As Collections.Generic.Dictionary(Of String, String) = Nothing

            If Not startLiteralLookup.TryGetValue(c, _value) Then
                Return Nothing
            End If

            Dim num As Integer = start
            start += 1
            Dim stringBuilder1 As New StringBuilder
            stringBuilder1.Append(c)
            Dim length1 As Integer = textBuffer.Length

            While start < length1
                stringBuilder1.Append(textBuffer(start))
                If _value.ContainsKey(stringBuilder1.ToString()) Then
                    start += 1
                    Continue While
                End If

                stringBuilder1.Length -= 1
                Exit While
            End While

            Return New Token(num, start - num, stringBuilder1.ToString(), Classification)
        End Function
    End Class
End Namespace
