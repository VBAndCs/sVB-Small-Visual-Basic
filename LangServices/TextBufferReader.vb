Imports System.IO
Imports Microsoft.Nautilus.Text

Namespace Microsoft.SmallVisualBasic.LanguageService
    Public Class TextBufferReader
        Inherits TextReader

        Private current As Integer
        Private textSnapshot As ITextSnapshot

        Public Sub New(textSnapshot As ITextSnapshot)
            Me.textSnapshot = textSnapshot
        End Sub

        Protected Overrides Sub Dispose(disposing As Boolean)
            MyBase.Dispose(disposing)
        End Sub

        Public Overrides Function Peek() As Integer
            If current = textSnapshot.Length Then
                Return -1
            End If

            Return AscW(textSnapshot(current))
        End Function

        Public Overloads Overrides Function Read() As Integer
            If current = textSnapshot.Length Then
                Return -1
            End If

            Read = AscW(textSnapshot(current))
            current += 1
        End Function

        Public Overloads Overrides Function Read(buffer As Char(), index As Integer, count As Integer) As Integer
            Dim i As Integer

            For i = 0 To count - 1

                If current >= textSnapshot.Length Then
                    Exit For
                End If

                buffer(index + i) = textSnapshot(current)
                current += 1
            Next

            Return i
        End Function

        Public Overrides Function ReadBlock(buffer As Char(), index As Integer, count As Integer) As Integer
            Dim i As Integer

            For i = 0 To count - 1

                If current >= textSnapshot.Length Then
                    Exit For
                End If

                buffer(index + i) = textSnapshot(current)
                current += 1
            Next

            Return i
        End Function

        Public Overrides Function ReadLine() As String
            If current >= textSnapshot.Length Then
                Return Nothing
            End If

            Dim line = textSnapshot.GetLineFromPosition(current)
            Dim text = line.GetText()

            If line.Start < current AndAlso line.End > current Then
                text = text.Substring(current - line.Start)
            End If

            current = line.EndIncludingLineBreak
            Return text
        End Function

        Public Overrides Function ReadToEnd() As String
            Dim text = textSnapshot.GetText(current, textSnapshot.Length - current)
            current = textSnapshot.Length
            Return text
        End Function
    End Class
End Namespace
