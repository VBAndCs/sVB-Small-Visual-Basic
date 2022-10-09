Imports System.Reflection.Emit

Namespace Microsoft.SmallVisualBasic.Statements
    Public MustInherit Class LoopStatement
        Inherits Statement

        Public Overridable Property ContinueLabel As Label
        Public Overridable Property ExitLabel As Label

        Public Body As New List(Of Statement)()
        Public EndLoopToken As Token
        Friend JumpLoopStatements As New List(Of JumpLoopStatement)

        Public Overrides Sub InferType(symbolTable As SymbolTable)
            For Each st In Body
                st.InferType(symbolTable)
            Next
        End Sub
    End Class
End Namespace
