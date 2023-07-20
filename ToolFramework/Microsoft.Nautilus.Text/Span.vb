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

            If length < 0 Then
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
            Return position >= _Start AndAlso position < [End]
        End Function

        Public Function Contains(span As Span) As Boolean
            Return span._Start >= _Start AndAlso span.End <= [End]
        End Function

        Public Function OverlapsWith(span As Span) As Boolean
            If span = Me Then Return True
            Return span._Start < [End] AndAlso span.End > _Start
        End Function

        Public Function Overlap(span As Span) As Span?
            If OverlapsWith(span) Then
                Dim start = Math.Max(_Start, span._Start)
                Dim [end] = Math.Min(Me.[End], span.End)
                Return New Span(start, [end] - start)
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
            If TypeOf obj IsNot Span Then Return False
            Dim span = CType(obj, Span)
            Return span._Start = _Start AndAlso span._Length = _Length
        End Function

        Public Shared Operator =(span1 As Span, span2 As Span) As Boolean
            Return span1._Start = span2._Start AndAlso span1._Length = span2._Length
        End Operator

        Public Shared Operator <>(span1 As Span, span2 As Span) As Boolean
            Return Not (span1 = span2)
        End Operator

    End Structure
End Namespace
