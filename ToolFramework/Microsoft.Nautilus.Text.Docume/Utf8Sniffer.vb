Imports System.ComponentModel.Composition
Imports System.IO
Imports System.Text
Imports Microsoft.Contracts
Imports Microsoft.Nautilus.Core

Namespace Microsoft.Nautilus.Text.Document

    <Export(GetType(IEncodingSniffer))>
    <Name("Utf8Sniffer")>
    Public Class Utf8Sniffer
        Implements IEncodingSniffer

        Public Function GetStreamEncoding(stream As Stream) As Encoding Implements IEncodingSniffer.GetStreamEncoding
            Contract.RequiresNotNull(stream, "stream")
            If Not IsUtf8(stream) Then Return Nothing
            Return Encoding.UTF8
        End Function

        Private Shared Function IsUtf8(s As Stream) As Boolean
            Dim array As Byte() = New Byte(1023) {}
            Dim flag As Boolean = False
            Dim num As Integer = 0
            Dim b As Byte = 128
            Dim b2 As Byte = 191

            While True
                Dim num2 As Integer = s.Read(array, 0, array.Length)
                If num2 = 0 Then Exit While

                For i As Integer = 0 To num2 - 1
                    Dim b3 As Byte = array(i)
                    If num = 0 Then
                        If b3 <= 127 Then Continue For

                        If b3 <= 223 Then
                            If b3 < 194 Then Return False

                            num = 1
                            Continue For
                        End If

                        If b3 <= 239 Then
                            num = 2

                            Select Case b3
                                Case 237
                                    b2 = 159
                                Case 224
                                    b = 160
                            End Select

                            Continue For
                        End If

                        If b3 > 244 Then Return False

                        num = 3
                        Select Case b3
                            Case 240
                                b = 144
                            Case 244
                                b2 = 143
                        End Select

                    Else
                        If b3 < b OrElse b3 > b2 Then
                            Return False
                        End If

                        If Threading.Interlocked.Decrement(num) = 0 Then
                            flag = True
                            Continue For
                        End If

                        b = 128
                        b2 = 191
                    End If
                Next
            End While

            If flag Then Return num = 0

            Return False

        End Function
    End Class
End Namespace
