Imports System
Imports Microsoft.SmallBasic.Statements

Namespace Microsoft.SmallBasic.Expressions
    <Serializable>
    Public MustInherit Class Expression
        Public Property StartToken As Token
        Public Property EndToken As Token
        Public Property Precedence As Integer

        Public Parent As Statement

        Public Overridable Sub AddSymbols(symbolTable As SymbolTable)
            _StartToken.Parent = Me.Parent
            _EndToken.Parent = Me.Parent
        End Sub

        Public Overridable Sub EmitIL(scope As CodeGenScope)
        End Sub
    End Class
End Namespace
