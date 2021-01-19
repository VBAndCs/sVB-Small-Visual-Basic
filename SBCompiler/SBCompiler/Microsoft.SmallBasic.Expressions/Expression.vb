Imports System

Namespace Microsoft.SmallBasic.Expressions
    <Serializable>
    Public MustInherit Class Expression
        Public Property StartToken As TokenInfo
        Public Property EndToken As TokenInfo
        Public Property Precedence As Integer

        Public Overridable Sub AddSymbols(ByVal symbolTable As SymbolTable)
        End Sub

        Public Overridable Sub EmitIL(ByVal scope As CodeGenScope)
        End Sub
    End Class
End Namespace
