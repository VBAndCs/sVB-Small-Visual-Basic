Imports Microsoft.Nautilus.Text

Namespace SuperClassifier
    Friend NotInheritable Class DelimitedLexExpression
        Inherits LexExpression

        Friend ReadOnly Start As String
        Friend ReadOnly [End] As String
        Friend ReadOnly IsMultiLine As Boolean
        Friend ReadOnly Ignore As String

        Friend Sub New(start As String, [end] As String, isMultiLine As Boolean, ignore As String, classification As String)
            MyBase.New(LexExpressionKind.Delimited, classification)
            Me.Start = start
            Me.End = [end]
            Me.IsMultiLine = isMultiLine
            Me.Ignore = ignore
        End Sub

        Friend Overrides Function GetTokenStartingAt(textBuffer As ITextSnapshot, start As Integer) As Token
            Dim length1 As Integer = textBuffer.Length
            If start + Me.Start.Length > length1 OrElse Not textBuffer.GetText(start, Me.Start.Length).Equals(Me.Start) Then
                Return Nothing
            End If

            Dim oldStart As Integer = start
            start += Me.Start.Length

            If [End] Is Nothing Then
                While start < length1
                    Dim c As Char = textBuffer(start)
                    If c = vbLf OrElse c = vbCr Then Exit While
                    start += 1
                End While
                Return New Token(oldStart, start - oldStart, textBuffer.GetText(oldStart, start - oldStart), Classification)
            End If

            If IsMultiLine Then
                Dim c2 As Char = [End](0)
                Dim length2 As Integer = [End].Length
                Dim c3 As Char = (If((Ignore IsNot Nothing), Ignore(0), ChrW(0)))
                Dim num2 As Integer = (If((Ignore IsNot Nothing), Ignore.Length, 0))

                While start < length1
                    Dim c4 As Char = textBuffer(start)
                    If c4 = c3 AndAlso num2 > 0 AndAlso start + num2 < length1 AndAlso textBuffer.GetText(start, num2).Equals(Ignore) Then
                        start += num2
                        Continue While
                    End If

                    If c4 = c2 AndAlso start + length2 < length1 AndAlso textBuffer.GetText(start, length2).Equals([End]) Then
                        start += length2
                        Exit While
                    End If
                    start += 1
                End While

                Return New Token(oldStart, start - oldStart, textBuffer.GetText(oldStart, start - oldStart), Classification)
            End If

            Dim c5 As Char = [End](0)
            Dim length3 As Integer = [End].Length
            Dim c6 As Char = (If((Ignore IsNot Nothing), Ignore(0), ChrW(0)))
            Dim num3 As Integer = (If((Ignore IsNot Nothing), Ignore.Length, 0))

            While start < length1
                Dim c7 As Char = textBuffer(start)
                If c7 = vbLf OrElse c7 = vbCr Then
                    Exit While
                End If

                If c7 = c6 AndAlso num3 > 0 AndAlso start + num3 < length1 AndAlso textBuffer.GetText(start, num3).Equals(Ignore) Then
                    start += num3
                    Continue While
                End If

                If c7 = c5 AndAlso start + length3 < length1 AndAlso textBuffer.GetText(start, length3).Equals([End]) Then
                    start += length3
                    Exit While
                End If

                start += 1
            End While

            Return New Token(oldStart, start - oldStart, textBuffer.GetText(oldStart, start - oldStart), Classification)
        End Function

    End Class
End Namespace
