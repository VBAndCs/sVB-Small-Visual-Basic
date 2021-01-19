Namespace Microsoft.SmallBasic.Statements
    Public Class EmptyStatement
        Inherits Statement

        Public Overrides Function ToString() As String
            Return VisualBasic.Constants.vbCrLf
        End Function
    End Class
End Namespace
