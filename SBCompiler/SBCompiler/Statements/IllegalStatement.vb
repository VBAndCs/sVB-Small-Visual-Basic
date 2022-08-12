Namespace Microsoft.SmallBasic.Statements
    Public Class IllegalStatement
        Inherits Statement

        Public Sub New(startToken As TokenInfo)
            Me.StartToken = startToken
        End Sub

        Public Overrides Function GetStatement(lineNumber) As Statement
            If lineNumber = StartToken.Line Then Return Me
            Return Nothing
        End Function


        Public Overrides Function ToString() As String
            Return "<<Illegal>>"
        End Function
    End Class
End Namespace
