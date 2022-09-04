Imports System.Reflection.Emit

Namespace Microsoft.SmallBasic.Statements
    Public MustInherit Class LoopStatement
        Inherits Statement

        Public Overridable Property ContinueLabel As Label
        Public Overridable Property ExitLabel As Label

        Public Body As New List(Of Statement)()
        Public EndLoopToken As Token
        Friend JumpLoopStatements As New List(Of JumpLoopStatement)

    End Class
End Namespace
