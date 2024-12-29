Imports System
Imports Microsoft.SmallVisualBasic.Statements

Namespace Microsoft.SmallVisualBasic.Expressions
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

        Public MustOverride Function InferType(symbolTable As SymbolTable) As VariableType

        Public MustOverride Function Evaluate(runner As Engine.ProgramRunner) As Library.Primitive

        Public MustOverride Function ToVB() As String

    End Class
End Namespace
