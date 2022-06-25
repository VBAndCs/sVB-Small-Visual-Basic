Imports System.Globalization

Namespace Microsoft.Nautilus.Text
    Public Structure Span

        Public ReadOnly Property Start As Integer

        Public ReadOnly Property Length As Integer

        Public ReadOnly Property [End] As Integer
            Get
                Return _Start + _Length
            End Get
        End Property

        Public ReadOnly Property IsEmpty As Boolean
            Get
                Return _Length = 0
            End Get
        End Property

        Public Sub New(start As Integer, length As Integer)
            If start < 0 Then
                Throw New ArgumentOutOfRangeException("start")
            End If

            If start + length < start Then
                Throw New ArgumentOutOfRangeException("length")
            End If

            _Start = start
            _Length = length
        End Sub

        Public Shared Function FromBounds(start As Integer, [end] As Integer) As Span
            If start < 0 Then
                Throw New ArgumentOutOfRangeException("start")
            End If

            If [end] < start Then
                Throw New ArgumentOutOfRangeException("end")
            End If

            Return New Span(start, [end] - start)
        End Function

        Public Function Contains(position As Integer) As Boolean
            If position >= _Start Then
                Return position < [End]
            End If
            Return False
        End Function

        Public Function Contains(span1 As Span) As Boolean
            If span1._Start >= _Start Then
                Return span1.End <= [End]
            End If
            Return False
        End Function

        Public Function OverlapsWith(span1 As Span) As Boolean
            If span1 = Me Then
                Return True
            End If

            If span1._Start < [End] Then
                Return span1.End > _Start
            End If
            Return False
        End Function

        Public Function Overlap(span1 As Span) As Span?
            If OverlapsWith(span1) Then
                Dim num As Integer = Math.Max(_Start, span1._Start)
                Dim num2 As Integer = Math.Min([End], span1.End)
                Return New Span(num, num2 - num)
            End If
            Return Nothing
        End Function

        Public Overrides Function ToString() As String
            Return String.Format(CultureInfo.InvariantCulture, "[{0}..{1})", New Object(1) {
                _Start,
                _Start + _Length
            })
        End Function

        Public Overrides Function GetHashCode() As Integer
            Return _Length Xor _Start
        End Function

        Public Overrides Function Equals(obj As Object) As Boolean
            If TypeOf obj Is Span Then
                Dim span1 = CType(obj, Span)
                If span1._Start = _Start Then
                    Return span1._Length = _Length
                End If
                Return False
            End If
            Return False
        End Function

        Public Shared Operator =(span1 As Span, span2 As Span) As Boolean
            If span1._Start = span2._Start Then
                Return span1._Length = span2._Length
            End If
            Return False
        End Operator

        Public Shared Operator <>(span1 As Span, span2 As Span) As Boolean
            Return Not (span1 = span2)
        End Operator
    End Structure
End Namespace
