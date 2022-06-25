Imports System.Text
Imports SuperClassifier
Friend Class BinaryNode

    Friend ReadOnly Property Token As Token

    Friend Property LeftChild As BinaryNode

    Friend Property RightChild As BinaryNode

    Public ReadOnly Property [String] As String
        Get
            Dim stringBuilder1 As New StringBuilder
            If LeftChild IsNot Nothing Then
                stringBuilder1.Append(LeftChild.String)
            End If

            stringBuilder1.AppendLine(Token.ToString())
            If RightChild IsNot Nothing Then
                stringBuilder1.Append(RightChild.String)
            End If

            Return stringBuilder1.ToString()
        End Get
    End Property

    Friend Sub New(token As Token)
        _Token = token
    End Sub
End Class
