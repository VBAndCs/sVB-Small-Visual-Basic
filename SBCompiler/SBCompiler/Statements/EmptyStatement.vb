Imports Microsoft.SmallVisualBasic.Engine

Namespace Microsoft.SmallVisualBasic.Statements
    Public Class EmptyStatement
        Inherits Statement

        Public Overrides Function GetStatementAt(lineNumber As Integer) As Statement
            If lineNumber = StartToken.Line Then Return Me
            Return Nothing
        End Function

        Public Overrides Function ToString() As String
            Return EndingComment.Text & vbCrLf
        End Function

        Public Overrides Function Execute(runner As ProgramRunner) As statement
            Return Nothing
        End Function

        Public Overrides Function ToVB(symbolTable As SymbolTable) As String
            Return EndingComment.Text & vbCrLf
        End Function
    End Class
End Namespace
