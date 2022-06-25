Imports System.Globalization

Namespace Microsoft.Nautilus.Text.StringRebuilder
    Public Structure LineSpan

        Private _span As Span

        Private _lineBreakLength As Integer

        Public ReadOnly Property LineNumber As Integer

        Public ReadOnly Property Start As Integer
            Get
                Return _span.Start
            End Get
        End Property


        Public ReadOnly Property [End] As Integer
            Get
                Return _span.End
            End Get
        End Property

        Public ReadOnly Property EndIncludingLineBreak As Integer
            Get
                Return _span.End + _lineBreakLength
            End Get
        End Property

        Public ReadOnly Property Length As Integer
            Get
                Return _span.Length
            End Get
        End Property

        Public ReadOnly Property LineBreakLength As Integer
            Get
                Return _lineBreakLength
            End Get
        End Property

        Public ReadOnly Property LengthIncludingLineBreak As Integer
            Get
                Return _span.Length + _lineBreakLength
            End Get
        End Property

        Public ReadOnly Property Extent As Span
            Get
                Return _span
            End Get
        End Property

        Public ReadOnly Property ExtentIncludingLineBreak As Span
            Get
                Return Span.FromBounds(_span.Start, EndIncludingLineBreak)
            End Get
        End Property

        Public Sub New(lineNumber As Integer, span As Span, lineBreakLength As Integer)
            If lineNumber < 0 Then
                Throw New ArgumentOutOfRangeException("lineNumber")
            End If

            If lineBreakLength > 2 OrElse span.End + lineBreakLength < span.End Then
                Throw New ArgumentOutOfRangeException("lineBreakLength")
            End If

            _LineNumber = lineNumber
            _span = span
            _lineBreakLength = lineBreakLength
        End Sub

        Public Overrides Function ToString() As String
            Return String.Format(CultureInfo.InvariantCulture, "[{0}, {1}+{2}]", New Object(2) {Start, [End], LineBreakLength})
        End Function

        Public Overrides Function Equals(obj As Object) As Boolean
            If TypeOf obj Is LineSpan Then
                Return CType(obj, LineSpan) = Me
            End If

            Return False
        End Function

        Public Shared Operator =(ls1 As LineSpan, ls2 As LineSpan) As Boolean
            If ls1._lineBreakLength = ls2._lineBreakLength AndAlso ls1._lineNumber = ls2._lineNumber Then
                Return ls1._span = ls2._span
            End If

            Return False
        End Operator

        Public Shared Operator <>(ls1 As LineSpan, ls2 As LineSpan) As Boolean
            Return Not (ls1 = ls2)
        End Operator

    End Structure
End Namespace
