Imports Microsoft.SmallBasic

Public Class LegalTokens
    Inherits List(Of TokenInfo)

    Public Overloads Sub Add(tokenInfo As TokenInfo)
        If tokenInfo.Token = Token.Illegal Then Return
        MyBase.Add(tokenInfo)
    End Sub
End Class
