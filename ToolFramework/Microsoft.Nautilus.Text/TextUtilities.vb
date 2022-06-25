Imports System.Text

Namespace Microsoft.Nautilus.Text
    Friend NotInheritable Class TextUtilities

        Public Shared Function ScanForLineCount(text As String) As Integer
            If text Is Nothing Then
                Throw New ArgumentNullException("text")
            End If

            Dim num As Integer = 0
            Dim num2 As Integer

            For num2 = text.Length - 1 To 0 + 1 Step -1
                Select Case text(num2).ToString()
                    Case vbLf
                        If text(num2 - 1) = vbCr Then
                            num2 -= 1
                        End If
                        num += 1

                    Case vbVerticalTab, vbFormFeed, vbCr, ChrW(&H85)
                        num += 1

                End Select
            Next

            If num2 = 0 Then
                Dim c As Char = text(0)
                If c = vbLf OrElse c = vbCr OrElse c = vbVerticalTab OrElse c = vbFormFeed OrElse c = ChrW(&H85) Then
                    num += 1
                End If
            End If

            Return num
        End Function

        Public Shared Function ScanForLineCount(chars As Char(), length1 As Integer) As Integer
            If chars Is Nothing Then
                Throw New ArgumentNullException("chars")
            End If

            If length1 < 0 Then
                Throw New ArgumentOutOfRangeException("length")
            End If

            Dim num As Integer = 0
            Dim num2 As Integer = 0

            For num2 = length1 - 1 To 0 + 1 Step -1
                Select Case chars(num2).ToString()
                    Case vbLf
                        If chars(num2 - 1) = vbCr Then
                            num2 -= 1
                        End If
                        num += 1

                    Case vbVerticalTab, vbFormFeed, vbCr, ChrW(&H85)
                        num += 1
                End Select

            Next

            If num2 = 0 Then
                Dim c As Char = chars(0)
                If c = vbLf OrElse c = vbCr OrElse c = vbVerticalTab OrElse c = vbFormFeed OrElse c = ChrW(&H85) Then
                    num += 1
                End If
            End If

            Return num
        End Function

        Public Shared Function ScanForLineCount(sb As StringBuilder) As Integer
            If sb Is Nothing Then
                Throw New ArgumentNullException("sb")
            End If

            Dim num As Integer = 0
            Dim num2 As Integer

            For num2 = sb.Length - 1 To 0 + 1 Step -1
                Select Case sb(num2).ToString()
                    Case vbLf
                        If sb(num2 - 1) = vbCr Then
                            num2 -= 1
                        End If
                        num += 1

                    Case vbVerticalTab, vbFormFeed, vbCr, ChrW(&H85)
                        num += 1

                End Select
            Next

            If num2 = 0 Then
                Dim c As Char = sb(0)
                If c = vbLf OrElse c = vbCr OrElse c = vbVerticalTab OrElse c = vbFormFeed OrElse c = ChrW(&H85) Then
                    num += 1
                End If
            End If

            Return num
        End Function
    End Class
End Namespace
