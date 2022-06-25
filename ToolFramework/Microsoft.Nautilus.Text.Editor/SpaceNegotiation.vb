Imports System.Windows

Namespace Microsoft.Nautilus.Text.Editor
    Public Structure SpaceNegotiation

        Public ReadOnly Property TextPosition As Integer

        Public ReadOnly Property Size As Size

        Public Sub New(textPosition As Integer, size As Size)
            If textPosition < 0 Then
                Throw New ArgumentOutOfRangeException("textPosition")
            End If

            If size.Width < 0.0 OrElse size.Height < 0.0 Then
                Throw New ArgumentOutOfRangeException("size")
            End If

            _TextPosition = textPosition
            _Size = size
        End Sub

        Public Overrides Function Equals(obj As Object) As Boolean
            If obj Is Nothing OrElse TypeOf obj IsNot SpaceNegotiation Then
                Return False
            End If

            Dim spaceNegotiation = CType(obj, SpaceNegotiation)
            If spaceNegotiation.Size = Size Then
                Return spaceNegotiation.TextPosition = TextPosition
            End If
            Return False
        End Function

        Public Overrides Function GetHashCode() As Integer
            Return _TextPosition.GetHashCode() Xor _Size.GetHashCode()
        End Function

        Public Shared Operator =(negotiation1 As SpaceNegotiation, negotiation2 As SpaceNegotiation) As Boolean
            Return negotiation1.Equals(negotiation2)
        End Operator

        Public Shared Operator <>(negotiation1 As SpaceNegotiation, negotiation2 As SpaceNegotiation) As Boolean
            Return Not negotiation1.Equals(negotiation2)
        End Operator
    End Structure
End Namespace
