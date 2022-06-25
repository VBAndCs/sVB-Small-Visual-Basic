Imports System.Globalization

Namespace Microsoft.Nautilus.Text.Editor
    Friend Class CaretPosition
        Implements ICaretPosition

        Public ReadOnly Property CharacterIndex As Integer Implements ICaretPosition.CharacterIndex

        Public ReadOnly Property Placement As CaretPlacement Implements ICaretPosition.Placement

        Public ReadOnly Property TextInsertionIndex As Integer Implements ICaretPosition.TextInsertionIndex

        Friend Sub New(characterIndex As Integer, textInsertionIndex As Integer, caretPlacement As CaretPlacement)
            If characterIndex < 0 Then
                Throw New ArgumentOutOfRangeException("characterIndex")
            End If

            If textInsertionIndex < 0 Then
                Throw New ArgumentOutOfRangeException("textInsertionIndex")
            End If

            _CharacterIndex = characterIndex
            _TextInsertionIndex = textInsertionIndex
            _Placement = caretPlacement
        End Sub

        Public Overrides Function ToString() As String
            If Placement = CaretPlacement.LeadingEdgeOfCharacter Then
                Return String.Format(CultureInfo.InvariantCulture, "|{0}", New Object(0) {CharacterIndex})
            End If

            Return String.Format(CultureInfo.InvariantCulture, "{0}|", New Object(0) {CharacterIndex})
        End Function

        Public Overrides Function GetHashCode() As Integer
            Return CharacterIndex.GetHashCode() * TextInsertionIndex.GetHashCode() * Placement.GetHashCode()
        End Function

        Public Overrides Function Equals(obj As Object) As Boolean
            If obj Is Nothing OrElse TypeOf obj IsNot CaretPosition Then Return False

            Dim caretPos = CType(obj, CaretPosition)
            Return caretPos.CharacterIndex = CharacterIndex AndAlso caretPos.TextInsertionIndex = TextInsertionIndex AndAlso caretPos.Placement = Placement

        End Function
    End Class
End Namespace
