Imports System.IO

Namespace Microsoft.Nautilus.Text
    Public NotInheritable Class TextSnapshotToTextReader
        Inherits TextReader

        Private _snapshot As ITextSnapshot
        Private _currentPosition As Integer
        Private _readLastLine As Boolean

        Public Overrides Sub Close()
            _currentPosition = -1
            MyBase.Close()
        End Sub

        Protected Overrides Sub Dispose(disposing As Boolean)
            _currentPosition = -1
            MyBase.Dispose(disposing)
        End Sub

        Public Overrides Function Peek() As Integer
            If _currentPosition = -1 Then
                Throw New ObjectDisposedException("TextSnapshotToTextReader")
            End If

            If _currentPosition <> _snapshot.Length Then
                Return AscW(_snapshot(_currentPosition))
            End If

            Return -1
        End Function

        Public Overrides Function Read() As Integer
            If _currentPosition = -1 Then
                Throw New ObjectDisposedException("TextSnapshotToTextReader")
            End If

            If _currentPosition <> _snapshot.Length Then
                Dim pos = _currentPosition
                _currentPosition += 1
                Return AscW(_snapshot(pos))
            End If
            Return -1
        End Function

        Public Overrides Function Read(buffer As Char(), index As Integer, count As Integer) As Integer
            If _currentPosition = -1 Then
                Throw New ObjectDisposedException("TextSnapshotToTextReader")
            End If
            If buffer Is Nothing Then
                Throw New ArgumentNullException("buffer")
            End If
            If index < 0 Then
                Throw New ArgumentOutOfRangeException("index")
            End If
            If count < 0 Then
                Throw New ArgumentOutOfRangeException("count")
            End If
            If index + count < 0 OrElse index + count > buffer.Length Then
                Throw New ArgumentOutOfRangeException("count")
            End If
            Dim num As Integer = Math.Min(_snapshot.Length - _currentPosition, count)
            _snapshot.CopyTo(_currentPosition, buffer, index, num)
            _currentPosition += num
            Return num
        End Function
        Public Overrides Function ReadBlock(buffer As Char(), index As Integer, count As Integer) As Integer
            Return Read(buffer, index, count)
        End Function
        Public Overrides Function ReadLine() As String
            If _currentPosition = -1 Then
                Throw New ObjectDisposedException("TextSnapshotToTextReader")
            End If
            If _readLastLine Then
                Return Nothing
            End If
            Dim lineFromPosition As ITextSnapshotLine = _snapshot.GetLineFromPosition(_currentPosition)
            Dim endIncludingLineBreak1 As Integer = lineFromPosition.EndIncludingLineBreak
            Dim text As String = _snapshot.GetText(_currentPosition, endIncludingLineBreak1 - _currentPosition)
            _currentPosition = endIncludingLineBreak1
            If lineFromPosition.LineBreakLength = 0 Then
                _readLastLine = True
            End If
            Return text
        End Function
        Public Overrides Function ReadToEnd() As String
            Dim text As String = _snapshot.GetText(_currentPosition, _snapshot.Length - _currentPosition)
            _currentPosition = _snapshot.Length
            Return text
        End Function

        Public Sub New(textSnapshot As ITextSnapshot)
            If textSnapshot Is Nothing Then
                Throw New ArgumentNullException("textSnapshot")
            End If
            _snapshot = textSnapshot
        End Sub
    End Class
End Namespace
