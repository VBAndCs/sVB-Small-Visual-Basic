Imports System.IO
Imports Microsoft.Nautilus.Text

Namespace Microsoft.SmallBasic.LanguageService
    Public Class TextBufferReader
        Inherits TextReader

        Private current As Integer
        Private textSnapshot As ITextSnapshot

        Public Sub New(ByVal textSnapshot As ITextSnapshot)
            Me.textSnapshot = textSnapshot
        End Sub

        Protected Overrides Sub Dispose(ByVal disposing As Boolean)
            MyBase.Dispose(disposing)
        End Sub

        Public Overrides Function Peek() As Integer
            If current = textSnapshot.Length Then
                Return -1
            End If

            Return AscW(textSnapshot(current))
        End Function

        Public Overrides Overloads Function Read() As Integer
            If current = textSnapshot.Length Then
                Return -1
            End If

            Read = AscW(textSnapshot(current))
            current += 1
        End Function

        Public Overrides Overloads Function Read(ByVal buffer As Char(), ByVal index As Integer, ByVal count As Integer) As Integer
            Dim i As Integer

            For i = 0 To count - 1

                If current >= textSnapshot.Length Then
                    Exit For
                End If

                buffer(index + i) = textSnapshot(Math.Min(Threading.Interlocked.Increment(current), current - 1))
            Next

            Return i
        End Function

        Public Overrides Function ReadBlock(ByVal buffer As Char(), ByVal index As Integer, ByVal count As Integer) As Integer
            Dim i As Integer

            For i = 0 To count - 1

                If current >= textSnapshot.Length Then
                    Exit For
                End If

                buffer(index + i) = textSnapshot(Math.Min(Threading.Interlocked.Increment(current), current - 1))
            Next

            Return i
        End Function

        Public Overrides Function ReadLine() As String
            If current >= textSnapshot.Length Then
                Return Nothing
            End If

            Dim lineFromPosition = textSnapshot.GetLineFromPosition(current)
            Dim text As String = lineFromPosition.GetText()

            If lineFromPosition.Start < current AndAlso lineFromPosition.End > current Then
                text = text.Substring(current - lineFromPosition.Start)
            End If

            current = lineFromPosition.EndIncludingLineBreak
            Return text
        End Function

        Public Overrides Function ReadToEnd() As String
            Dim text = textSnapshot.GetText(current, textSnapshot.Length - current)
            current = textSnapshot.Length
            Return text
        End Function
    End Class
End Namespace
