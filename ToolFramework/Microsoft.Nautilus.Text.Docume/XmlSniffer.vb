Imports System.ComponentModel.Composition
Imports System.IO
Imports System.Text
Imports Microsoft.Nautilus.Core

Namespace Microsoft.Nautilus.Text.Document.Implementation

    <Export(GetType(IEncodingSniffer))>
    <OrderBefore("Utf8Sniffer")>
    <Name("XmlSniffer")>
    Public Class XmlSniffer
        Implements IEncodingSniffer

        Public Function GetStreamEncoding(stream1 As Stream) As Encoding Implements IEncodingSniffer.GetStreamEncoding
            If stream1 Is Nothing Then
                Throw New ArgumentNullException("stream")
            End If

            Dim position1 As Long = stream1.Position
            Dim array As Encoding() = New Encoding(2) {
                Encoding.ASCII,
                Encoding.Unicode,
                Encoding.BigEndianUnicode
            }

            For Each encoding1 As Encoding In array
                Dim reader As New StreamReader(stream1, encoding1, detectEncodingFromByteOrderMarks:=False)
                Dim encoding2 As Encoding = DetectXmlEncoding(reader)
                stream1.Position = position1
                If encoding2 IsNot Nothing Then
                    Return If((encoding1 Is Encoding.ASCII), encoding2, encoding1)
                End If
            Next
            Return Nothing
        End Function

        Private Shared Function DetectXmlEncoding(reader As TextReader) As Encoding
            Dim text As String = ExtractEncoding(reader)
            If text Is Nothing Then Return Nothing

            If String.Compare("unicode", text, StringComparison.OrdinalIgnoreCase) = 0 Then
                Return Encoding.Unicode
            End If

            Dim encodings As EncodingInfo() = Encoding.GetEncodings()
            Dim array As EncodingInfo() = encodings

            For Each encodingInfo1 As EncodingInfo In array
                If String.Compare(encodingInfo1.Name, text, StringComparison.OrdinalIgnoreCase) = 0 Then
                    Return encodingInfo1.GetEncoding()
                End If
            Next

            Return Nothing
        End Function

        Private Shared Function ExtractEncoding(reader As TextReader) As String
            If Not SkipXmlPrefix(reader) Then
                Return Nothing
            End If

            Dim flag As Boolean = False
            Dim text As String = ""
            Dim text2 As String = ""

            While True
                text = ExtractName(reader)
                If text Is Nothing Then
                    Return Nothing
                End If

                text2 = ExtractValue(reader)
                If text2 Is Nothing Then
                    Return Nothing
                End If

                If flag Then Exit While

                If String.Compare(text, "version", StringComparison.Ordinal) = 0 Then
                    flag = True
                    Continue While
                End If
                Return Nothing
            End While

            If String.Compare(text, "encoding", StringComparison.Ordinal) = 0 Then
                Return text2
            End If

            Return Nothing
        End Function

        Private Shared Function SkipXmlPrefix(reader As TextReader) As Boolean
            Dim num As Integer = SkipWhiteSpace(reader)
            If num = -1 Then
                Return False
            End If
            Dim text As String = "<?xml "
            Dim num2 As Integer = 0

            Do

                If Char.ToLowerInvariant(ChrW(num)) <> text(Math.Min(Threading.Interlocked.Increment(num2), num2 - 1)) Then
                    Return False
                End If
                If num2 = text.Length Then
                    Return True
                End If
                num = reader.Read()
            Loop While num <> -1
            Return False
        End Function

        Private Shared Function SkipWhiteSpace(reader As TextReader) As Integer
            Dim num As Integer

            Do
                num = reader.Read()
                If num = -1 Then
                    Return -1
                End If
            Loop While Char.IsWhiteSpace(ChrW(num))
            Return num
        End Function

        Private Shared Function ExtractName(reader As TextReader) As String
            Dim num As Integer = SkipWhiteSpace(reader)
            If num = -1 Then
                Return Nothing
            End If
            Dim stringBuilder1 As New StringBuilder
            While num <> 61 AndAlso Not Char.IsWhiteSpace(ChrW(num))
                If stringBuilder1.Length >= 8 Then
                    Return Nothing
                End If
                stringBuilder1.Append(ChrW(num))
                num = reader.Read()
                If num = -1 Then
                    Return Nothing
                End If
            End While
            If num <> 61 Then
                num = SkipWhiteSpace(reader)
                If num <> 61 Then
                    Return Nothing
                End If
            End If
            If stringBuilder1.Length <> 0 Then
                Return stringBuilder1.ToString()
            End If
            Return Nothing
        End Function

        Private Shared Function ExtractValue(reader As TextReader) As String
            Dim num As Integer = SkipWhiteSpace(reader)
            Select Case num
                Case 34, 39
                    Dim stringBuilder1 As New StringBuilder
                    While True
                        Dim num2 As Integer = reader.Read()
                        If num2 = -1 Then
                            Return Nothing
                        End If
                        If num2 = num Then
                            Exit While
                        End If
                        If stringBuilder1.Length >= 16 Then
                            Return Nothing
                        End If
                        stringBuilder1.Append(ChrW(num2))
                    End While
                    If stringBuilder1.Length <> 0 Then
                        Return stringBuilder1.ToString()
                    End If
                    Return Nothing
                Case Else
                    Return Nothing
            End Select
        End Function
    End Class
End Namespace
