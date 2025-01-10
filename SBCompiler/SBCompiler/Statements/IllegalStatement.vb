Imports Microsoft.SmallVisualBasic.Engine

Namespace Microsoft.SmallVisualBasic.Statements
    Public Class IllegalStatement
        Inherits Statement

        Public Sub New(startToken As Token)
            Me.StartToken = startToken
        End Sub

        Public Overrides Function GetStatementAt(lineNumber As Integer) As Statement
            If lineNumber = StartToken.Line Then Return Me
            Return Nothing
        End Function


        Public Overrides Function ToString() As String
            Return "<<Illegal>>" & vbCrLf
        End Function

        Public Overrides Function Execute(runner As ProgramRunner) As statement
            Return Nothing
        End Function

        Public Overrides Function ToVB(symbolTable As SymbolTable) As String
            Throw New NotImplementedException()
        End Function
    End Class
End Namespace
