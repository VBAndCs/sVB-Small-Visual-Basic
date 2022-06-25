Imports System.Collections
Imports System.Collections.Generic
Imports System.Globalization
Imports System.IO

Namespace Microsoft.Nautilus.Text.StringRebuilder
    Friend MustInherit Class BaseStringRebuilder
        Implements IStringRebuilder
        Implements IComparable(Of IStringRebuilder), IEnumerable(Of Char), IEnumerable, IEquatable(Of IStringRebuilder)

        Private Class RebuilderEnumerator
            Implements IEnumerator(Of Char), IDisposable, IEnumerator

            Private _index As Integer = -1
            Private _source As IStringRebuilder

            Public ReadOnly Property Current As Char Implements Collections.Generic.IEnumerator(Of Char).Current
                Get
                    If _index < 0 OrElse _index >= _source.Length Then
                        Throw New InvalidOperationException("Enumerator not at a valid position")
                    End If

                    Return _source(_index)
                End Get
            End Property

            ReadOnly Property IEnumerator_Current As Object Implements IEnumerator.Current
                Get
                    If _index < 0 OrElse _index >= _source.Length Then
                        Throw New InvalidOperationException("Enumerator not at a valid position")
                    End If

                    Return _source(_index)
                End Get
            End Property

            Public Sub New(source As IStringRebuilder)
                _source = source
            End Sub

            Public Sub Dispose() Implements IDisposable.Dispose
            End Sub

            Public Function MoveNext() As Boolean Implements Collections.IEnumerator.MoveNext
                Return Threading.Interlocked.Increment(_index) < _source.Length
            End Function

            Public Sub Reset() Implements Collections.IEnumerator.Reset
                _index = -1
            End Sub
        End Class

        Public MustOverride ReadOnly Property Length As Integer Implements IStringRebuilder.Length

        Public MustOverride ReadOnly Property LineBreakCount As Integer Implements IStringRebuilder.LineBreakCount

        Default Public MustOverride ReadOnly Property Item(index As Integer) As Char Implements IStringRebuilder.Item

        Private Function Assemble(left As Span, right As Span) As IStringRebuilder
            If left.Length = 0 Then Return Substring(right)

            If right.Length = 0 Then Return Substring(left)

            If left.Length + right.Length = Length Then Return Me

            Return BinaryStringRebuilder.Create(Substring(left), Substring(right))
        End Function

        Private Function Assemble(left As Span, text As IStringRebuilder, right As Span) As IStringRebuilder
            If text.Length = 0 Then
                Return Assemble(left, right)
            End If

            If left.Length = 0 Then
                If right.Length <> 0 Then
                    Return BinaryStringRebuilder.Create(text, Substring(right))
                End If

                Return text
            End If

            If right.Length = 0 Then
                Return BinaryStringRebuilder.Create(Substring(left), text)
            End If

            If left.Length < right.Length Then
                Return BinaryStringRebuilder.Create(BinaryStringRebuilder.Create(Substring(left), text), Substring(right))
            End If

            Return BinaryStringRebuilder.Create(Substring(left), BinaryStringRebuilder.Create(text, Substring(right)))
        End Function

        Public MustOverride Function GetLineNumberFromPosition(position As Integer) As Integer Implements IStringRebuilder.GetLineNumberFromPosition
        Public MustOverride Function GetLineFromLineNumber(lineNumber As Integer) As LineSpan Implements IStringRebuilder.GetLineFromLineNumber
        Public MustOverride Function CompareTo(other As IStringRebuilder, ignoreCase As Boolean, culture As CultureInfo) As Integer Implements IStringRebuilder.CompareTo
        Public MustOverride Function GetText(span1 As Span) As String Implements IStringRebuilder.GetText

        Public Function ToCharArray(startIndex As Integer, length1 As Integer) As Char() Implements IStringRebuilder.ToCharArray
            If startIndex < 0 Then
                Throw New ArgumentOutOfRangeException("startIndex")
            End If

            If length1 < 0 OrElse startIndex + length1 > Length OrElse startIndex + length1 < 0 Then
                Throw New ArgumentOutOfRangeException("length")
            End If

            Dim array As Char() = New Char(length1 - 1) {}
            CopyTo(startIndex, array, 0, length1)
            Return array
        End Function

        Public MustOverride Sub CopyTo(sourceIndex As Integer, destination As Char(), destinationIndex As Integer, count As Integer) Implements IStringRebuilder.CopyTo
        Public MustOverride Sub Write(writer As TextWriter, span1 As Span) Implements IStringRebuilder.Write
        Public MustOverride Function Substring(span1 As Span) As IStringRebuilder Implements IStringRebuilder.Substring

        Public Function Delete(span1 As Span) As IStringRebuilder Implements IStringRebuilder.Delete
            If span1.End > Length Then
                Throw New ArgumentOutOfRangeException("span")
            End If

            Return Assemble(Span.FromBounds(0, span1.Start), Span.FromBounds(span1.End, Length))
        End Function

        Public Function Replace(span1 As Span, text As String) As IStringRebuilder Implements IStringRebuilder.Replace
            Return Replace(span1, SimpleStringRebuilder.Create(text))
        End Function

        Public Function Replace(span1 As Span, text As IStringRebuilder) As IStringRebuilder Implements IStringRebuilder.Replace
            If span1.End > Length Then
                Throw New ArgumentOutOfRangeException("span")
            End If

            If text Is Nothing Then
                Throw New ArgumentNullException("text")
            End If

            Return Assemble(Span.FromBounds(0, span1.Start), text, Span.FromBounds(span1.End, Length))
        End Function

        Public Function Insert(position As Integer, text As String) As IStringRebuilder Implements IStringRebuilder.Insert
            Return Insert(position, SimpleStringRebuilder.Create(text))
        End Function

        Public Function Insert(position As Integer, text As IStringRebuilder) As IStringRebuilder Implements IStringRebuilder.Insert
            If position < 0 OrElse position > Length Then
                Throw New ArgumentOutOfRangeException("position")
            End If

            If text Is Nothing Then
                Throw New ArgumentNullException("text")
            End If

            Return Assemble(Span.FromBounds(0, position), text, Span.FromBounds(position, Length))
        End Function

        Public Function CompareTo(other As IStringRebuilder) As Integer Implements IComparable(Of IStringRebuilder).CompareTo
            Return CompareTo(other, ignoreCase:=False, CultureInfo.InvariantCulture)
        End Function

        Public Function GetEnumerator() As IEnumerator(Of Char) Implements Collections.Generic.IEnumerable(Of Char).GetEnumerator
            Return New RebuilderEnumerator(Me)
        End Function

        Private Function GetEnumerator1() As IEnumerator Implements Collections.IEnumerable.GetEnumerator
            Return New RebuilderEnumerator(Me)
        End Function

        Public Overloads Function Equals(other As IStringRebuilder) As Boolean Implements IEquatable(Of IStringRebuilder).Equals
            Return CompareTo(other) = 0
        End Function
    End Class
End Namespace
