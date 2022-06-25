Imports System.Globalization

Namespace Microsoft.Nautilus.Text.Editor
    Public Structure TextBounds

        Public ReadOnly Property Left As Double

        Public ReadOnly Property Top As Double

        Public ReadOnly Property Width As Double

        Public ReadOnly Property Height As Double

        Public ReadOnly Property Right As Double
            Get
                Return _Left + _Width
            End Get
        End Property

        Public ReadOnly Property Bottom As Double
            Get
                Return _Top + _Height
            End Get
        End Property

        Public Sub New(left As Double, top As Double, width As Double, height As Double)

            If Double.IsNaN(left) Then
                Throw New ArgumentOutOfRangeException("left")
            End If

            If Double.IsNaN(top) Then
                Throw New ArgumentOutOfRangeException("top")
            End If

            If Double.IsNaN(width) Then
                Throw New ArgumentOutOfRangeException("width")
            End If

            If Double.IsNaN(height) OrElse height < 0.0 Then
                Throw New ArgumentOutOfRangeException("height")
            End If

            _Left = left
            _Top = top
            _Width = width
            _Height = height

        End Sub

        Public Overrides Function ToString() As String
            Return String.Format(CultureInfo.InvariantCulture, "[{0},{1},{2},{3}]", Left, Top, Right, Bottom)
        End Function

        Public Overrides Function GetHashCode() As Integer
            Return _Left.GetHashCode() Xor _Top.GetHashCode() Xor _Width.GetHashCode() Xor _Height.GetHashCode()
        End Function

        Public Overrides Function Equals(obj As Object) As Boolean
            If obj Is Nothing OrElse TypeOf obj IsNot TextBounds Then
                Return False
            End If

            Return CType(obj, TextBounds) = Me
        End Function

        Public Shared Operator =(bounds1 As TextBounds, bounds2 As TextBounds) As Boolean
            If bounds1.Left = bounds2.Left AndAlso bounds1.Top = bounds2.Top AndAlso bounds1.Width = bounds2.Width Then
                Return bounds1.Height = bounds2.Height
            End If
            Return False
        End Operator

        Public Shared Operator <>(bounds1 As TextBounds, bounds2 As TextBounds) As Boolean
            Return Not (bounds1 = bounds2)
        End Operator
    End Structure
End Namespace
