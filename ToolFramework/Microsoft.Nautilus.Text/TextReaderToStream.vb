Imports System.IO
Imports System.Text

Namespace Microsoft.Nautilus.Text
    Public NotInheritable Class TextReaderToStream
        Inherits Stream

        Private _absolutePosition As Long
        Private _reader As TextReader
        Private _writer As StreamWriter
        Private _memoryStream As MemoryStream
        Private _buffer As Byte() = New Byte(1023) {}
        Private _bufferPosition As Integer
        Private _bytesInBuffer As Integer
        Private _cachedSurrogate As Char
        Private disposed As Boolean

        Public Overrides ReadOnly Property CanRead As Boolean = True

        Public Overrides ReadOnly Property CanSeek As Boolean = False

        Public Overrides ReadOnly Property CanWrite As Boolean = False

        Public Overrides ReadOnly Property Length As Long = -1

        Public Overrides Property Position As Long
            Get
                Return _absolutePosition
            End Get

            Set(value As Long)
                Throw New NotSupportedException
            End Set
        End Property

        Public Overrides Sub Flush()
            Throw New NotSupportedException
        End Sub

        Public Overrides Function Read(buffer As Byte(), offset As Integer, count As Integer) As Integer
            If disposed Then
                Throw New ObjectDisposedException("TextReaderToStream")
            End If

            If buffer Is Nothing Then
                Throw New ArgumentNullException("buffer")
            End If

            If offset < 0 Then
                Throw New ArgumentOutOfRangeException("offset")
            End If

            If count < 0 Then
                Throw New ArgumentOutOfRangeException("count")
            End If

            If offset + count < 0 OrElse offset + count > buffer.Length Then
                Throw New ArgumentOutOfRangeException("count")
            End If

            Dim num As Integer = 0
            While count > 0
                Dim num2 As Integer = Math.Min(count, _bytesInBuffer - _bufferPosition)
                If num2 > 0 Then
                    Array.Copy(_buffer, _bufferPosition, buffer, offset, num2)
                    num += num2
                    _bufferPosition += num2
                    offset += num2
                    count -= num2
                End If

                If count > 0 Then
                    Dim num3 As Integer = ReadBlock()
                    If num3 <= 0 Then
                        Exit While
                    End If
                End If
            End While

            _absolutePosition += num
            Return num
        End Function

        Public Overrides Function Seek(offset As Long, origin As SeekOrigin) As Long
            Throw New NotSupportedException
        End Function

        Public Overrides Sub SetLength(_value As Long)
            Throw New NotSupportedException
        End Sub

        Public Overrides Sub Write(buffer As Byte(), offset As Integer, count As Integer)
            Throw New NotSupportedException
        End Sub

        Protected Overrides Sub Dispose(disposing As Boolean)
            Try
                If Not disposed AndAlso disposing Then
                    _reader.Dispose()
                    _writer.Dispose()
                    _memoryStream.Dispose()
                    _reader = Nothing
                    _writer = Nothing
                    _memoryStream = Nothing
                    disposed = True
                End If

            Finally
                MyBase.Dispose(disposing)
            End Try
        End Sub

        Public Sub New(reader As TextReader, encoding1 As Encoding)
            If reader Is Nothing Then
                Throw New ArgumentNullException("reader")
            End If

            If encoding1 Is Nothing Then
                Throw New ArgumentNullException("encoding")
            End If

            _reader = reader
            _memoryStream = New MemoryStream(_buffer, writable:=True)
            _writer = New StreamWriter(_memoryStream, encoding1, _buffer.Length)

        End Sub

        Private Function ReadBlock() As Integer
            Dim block((_buffer.Length - 4) \ 4 - 1) As Char
            Dim index As Integer = 0

            If _cachedSurrogate <> vbNullChar Then
                block(0) = _cachedSurrogate
                index = 1
                _cachedSurrogate = ChrW(0)
            End If

            Dim count = _reader.Read(block, index, block.Length - index)
            count += index

            If count > 0 Then
                If count >= 2 AndAlso Char.IsHighSurrogate(block(count - 1)) Then
                    count -= 1
                    _cachedSurrogate = block(count)
                End If

                _memoryStream.Position = 0L
                _bufferPosition = 0
                _writer.Write(block, 0, count)
                _writer.Flush()
                _bytesInBuffer = CInt(CLng(Fix(_memoryStream.Position)) Mod Integer.MaxValue)
            End If

            Return count
        End Function
    End Class
End Namespace
