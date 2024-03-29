Imports Microsoft.SmallVisualBasic.Engine

Namespace Microsoft.SmallVisualBasic.Statements
    Public Class StopStatement
        Inherits Statement

        Public Overrides Function GetStatementAt(lineNumber As Integer) As Statement
            If lineNumber = StartToken.Line Then Return Me
            Return Nothing
        End Function

        Public Overrides Function ToString() As String
            Return StartToken.Text & vbCrLf
        End Function

        Public Overrides Function Execute(runner As ProgramRunner) As Statement
            Return Nothing
        End Function
    End Class
End Namespace

