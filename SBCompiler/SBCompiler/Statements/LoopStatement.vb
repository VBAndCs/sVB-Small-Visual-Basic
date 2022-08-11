Imports System.Reflection.Emit

Namespace Microsoft.SmallBasic.Statements
    Public MustInherit Class LoopStatement
        Inherits Statement

        Public ContinueLabel As Label
        Public ExitLabel As Label
        Public Body As New List(Of Statement)()
        Public EndLoopToken As TokenInfo

    End Class
End Namespace
