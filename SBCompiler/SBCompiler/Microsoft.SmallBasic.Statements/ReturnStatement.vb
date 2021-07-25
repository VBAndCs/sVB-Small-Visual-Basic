Imports System.Globalization
Imports Microsoft.SmallBasic.Expressions

Namespace Microsoft.SmallBasic.Statements
    Public Class ReturnStatement
        Inherits Statement

        Public ReturnExpression As Expression

        Public Overrides Function ToString() As String
            Return $"{StartToken} {ReturnExpression}"
        End Function
    End Class
End Namespace
