Imports System.Collections.Generic
Imports System.ComponentModel.Composition
Imports System.Globalization
Imports System.IO
Imports System.Text
Imports System.Threading
Imports Microsoft.Nautilus.Core.Task
Imports Microsoft.Nautilus.Text.Projection
Imports Microsoft.Nautilus.Text.Projection.Implementation
Imports Microsoft.Nautilus.Text.StringRebuilder

Namespace Microsoft.Nautilus.Text
    <Export(GetType(ITextBufferFactory))>
    <Export(GetType(IProjectionBufferFactory))>
    Public NotInheritable Class BufferFactory
        Implements ITextBufferFactory, IProjectionBufferFactory, IBufferTracker

        Private _trackBuffers As Boolean
        Private allocatedBuffers As List(Of WeakReference)

        Public Property TrackBuffers As Boolean Implements IBufferTracker.TrackBuffers
            Get
                Return _trackBuffers
            End Get

            Set(value As Boolean)
                _trackBuffers = value
                If value AndAlso allocatedBuffers Is Nothing Then
                    allocatedBuffers = New List(Of WeakReference)
                End If
            End Set
        End Property

        Public Sub TagBuffer(buffer As ITextBuffer, tag As String) Implements IBufferTracker.TagBuffer
            If buffer Is Nothing Then
                Throw New ArgumentNullException("buffer")
            End If

            If tag Is Nothing Then
                Throw New ArgumentNullException("tag")
            End If

            buffer.Properties.AddProperty("tag", tag)
        End Sub

        Public Function ReportLiveBuffers(writer As TextWriter) As Integer Implements IBufferTracker.ReportLiveBuffers
            Dim num As Integer = 0
            If allocatedBuffers IsNot Nothing Then
                For i As Integer = 0 To allocatedBuffers.Count - 1

                    If TypeOf allocatedBuffers(i).Target IsNot ITextBuffer Then
                        Continue For
                    End If

                    Dim textBuffer = CType(allocatedBuffers(i).Target, ITextBuffer)

                    num += 1
                    If writer Is Nothing Then
                        Continue For
                    End If

                    Dim [property] As String = ""
                    Dim flag As Boolean = textBuffer.Properties.TryGetProperty(Of String)("tag", [property])
                    writer.Write(String.Format(
                        CultureInfo.CurrentCulture,
                        "{0,5} {1}" & vbCrLf,
                        New Object() {i, If(flag, [property], "Untagged")})
                    )

                    For Each property2 As KeyValuePair(Of Object, Object) In textBuffer.Properties
                        If Not (property2.Key.ToString() <> "tag") Then
                            Continue For
                        End If

                        Dim text As String
                        If property2.Value Is Nothing Then
                            text = "?null"
                        Else
                            text = property2.Value.GetType().Name
                            If TypeOf property2.Value Is WeakReference Then
                                Dim target1 As Object = TryCast(property2.Value, WeakReference).Target
                                If target1 IsNot Nothing Then
                                    text = text & "(" & target1.GetType().Name & ")"
                                End If
                            End If
                        End If
                        writer.Write(String.Format(CultureInfo.CurrentCulture, "      {0}" & vbCrLf, New Object(0) {text}))
                    Next
                Next
            End If
            Return num
        End Function

        Private Sub Record(buffer As ITextBuffer)
            For i As Integer = 0 To allocatedBuffers.Count - 1
                If Not allocatedBuffers(i).IsAlive Then
                    allocatedBuffers(i).Target = buffer
                    Return
                End If
            Next
            allocatedBuffers.Add(New WeakReference(buffer))
        End Sub

        Private Function Make(contentType As String, content As IStringRebuilder) As TextBuffer
            Dim textBuffer As New TextBuffer(contentType, content)
            If _trackBuffers Then
                Record(textBuffer)
            End If
            Return textBuffer
        End Function

        Public Function CreateTextBuffer() As ITextBuffer Implements ITextBufferFactory.CreateTextBuffer
            Return CreateTextBuffer("")
        End Function

        Public Function CreateTextBuffer(text As String) As ITextBuffer Implements ITextBufferFactory.CreateTextBuffer
            If text Is Nothing Then
                Throw New ArgumentNullException("text")
            End If
            Return CreateTextBuffer(text, "text")
        End Function

        Public Function CreateTextBuffer(text As String, contentType As String) As ITextBuffer Implements ITextBufferFactory.CreateTextBuffer
            If text Is Nothing Then
                Throw New ArgumentNullException("text")
            End If

            If contentType Is Nothing Then
                Throw New ArgumentNullException("contentType")
            End If

            Return Make(contentType, SimpleStringRebuilder.Create(text))
        End Function

        Public Function CreateTextBuffer(reader As TextReader) As ITextBuffer Implements ITextBufferFactory.CreateTextBuffer
            If reader Is Nothing Then
                Throw New ArgumentNullException("reader")
            End If

            Return CreateTextBuffer(reader, "text")
        End Function

        Public Function CreateTextBuffer(reader As TextReader, contentType As String) As ITextBuffer Implements ITextBufferFactory.CreateTextBuffer
            If reader Is Nothing Then
                Throw New ArgumentNullException("reader")
            End If

            If contentType Is Nothing Then
                Throw New ArgumentNullException("contentType")
            End If

            Dim stringRebuilder As IStringRebuilder = SimpleStringRebuilder.Create("")
            Dim array As Char() = New Char(4095) {}
            Dim num As Integer

            Do
                num = reader.ReadBlock(array, 0, array.Length)
                Dim stringBuilder1 As New StringBuilder(num)
                stringBuilder1.Append(array, 0, num)
                stringRebuilder = stringRebuilder.Insert(stringRebuilder.Length, stringBuilder1.ToString())
            Loop While num >= array.Length

            Return Make(contentType, stringRebuilder)
        End Function

        Public Function CreateTextBuffer(reader As TextReader, contentType As String, cancellableTask As CancelableTask) As ITextBuffer Implements ITextBufferFactory.CreateTextBuffer
            If reader Is Nothing Then
                Throw New ArgumentNullException("reader")
            End If

            If contentType Is Nothing Then
                Throw New ArgumentNullException("contentType")
            End If

            If cancellableTask Is Nothing Then
                Throw New ArgumentNullException("cancellableTask")
            End If

            Dim flag As Boolean = False
            Dim stringRebuilder As IStringRebuilder = SimpleStringRebuilder.Create("")
            Dim array As Char() = New Char(4095) {}

            While True
                Dim num As Integer = reader.ReadBlock(array, 0, array.Length)
                Dim stringBuilder1 As New StringBuilder(num)
                stringBuilder1.Append(array, 0, num)
                stringRebuilder = stringRebuilder.Insert(stringRebuilder.Length, stringBuilder1.ToString())

                If num < array.Length Then Exit While
                If cancellableTask.CancelRequested Then Return Nothing
                If flag Then Thread.Sleep(200)

            End While

            Return Make(contentType, stringRebuilder)
        End Function

        Public Function CreateProjectionBuffer(projectionEditResolver As IProjectionEditResolver, contentType As String, textSpans As IList(Of ITextSpan)) As IProjectionBuffer Implements IProjectionBufferFactory.CreateProjectionBuffer
            Return New ProjectionBuffer(projectionEditResolver, contentType, textSpans)
        End Function
    End Class
End Namespace
