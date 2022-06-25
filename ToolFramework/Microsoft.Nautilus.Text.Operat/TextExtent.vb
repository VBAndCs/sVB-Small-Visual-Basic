Namespace Microsoft.Nautilus.Text.Operations
    Public Structure TextExtent

        Public ReadOnly Property Span As Span

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

        Public ReadOnly Property Length As Integer
            Get
                Return _span.Length
            End Get
        End Property

        Public ReadOnly Property IsSignificant As Boolean

        Public Sub New(start As Integer, length As Integer, isSignificant As Boolean)
            _span = New Span(start, length)
            _isSignificant = isSignificant
        End Sub

        Public Sub New(span As Span, isSignificant As Boolean)
            _Span = span
            _IsSignificant = isSignificant
        End Sub

        Public Sub New(textExtent As TextExtent)
            _Span = textExtent.Span
            _IsSignificant = textExtent.IsSignificant
        End Sub

        Public Overrides Function Equals(obj As Object) As Boolean
            If obj IsNot Nothing AndAlso TypeOf obj Is TextExtent Then
                Return Me = CType(obj, TextExtent)
            End If

            Return False
        End Function

        Public Shared Operator =(extent1 As TextExtent, extent2 As TextExtent) As Boolean
            If extent1._span = extent2._span Then
                Return extent1._isSignificant = extent2._isSignificant
            End If

            Return False
        End Operator

        Public Shared Operator <>(extent1 As TextExtent, extent2 As TextExtent) As Boolean
            Return Not (extent1 = extent2)
        End Operator
    End Structure
End Namespace
