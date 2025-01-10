Imports Microsoft.SmallVisualBasic.Engine

Namespace Microsoft.SmallVisualBasic.Statements
    ''' <summary>
    ''' This is not an actual a statement but just a flag to terminate the debugger
    ''' </summary>
    Public Class EndDebugging
        Inherits Statement

        Public Overrides Function GetStatementAt(lineNumber As Integer) As Statement
            Throw New NotImplementedException()
        End Function

        Public Overrides Function Execute(runner As ProgramRunner) As Statement
            Return Me
        End Function

        Public Overrides Function ToVB(symbolTable As SymbolTable) As String
            Return ""
        End Function
    End Class

End Namespace
