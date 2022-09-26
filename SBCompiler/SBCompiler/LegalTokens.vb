Imports Microsoft.SmallVisualBasic

Public Class LegalTokens
    Inherits List(Of Token)

    Public Overloads Sub Add(token As Token)
        If token.Type = TokenType.Illegal Then Return
        MyBase.Add(token)
    End Sub
End Class
