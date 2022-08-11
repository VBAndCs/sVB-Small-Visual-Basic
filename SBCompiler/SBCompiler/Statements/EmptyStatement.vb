Namespace Microsoft.SmallBasic.Statements
    Public Class EmptyStatement
        Inherits Statement

        Public Overrides Function GetStatement(lineNumber) As Statement
            If lineNumber = StartToken.Line Then Return Me
            Return Nothing
        End Function

        Public Overrides Function ToString() As String
            Return VisualBasic.Constants.vbCrLf
        End Function
    End Class
End Namespace
