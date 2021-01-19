Imports System
Imports System.Reflection.Emit

Namespace Microsoft.SmallBasic.Expressions
    <Serializable>
    Public Class IdentifierExpression
        Inherits Expression

        Public Property Identifier As TokenInfo

        Public Overrides Sub AddSymbols(ByVal symbolTable As SymbolTable)
            symbolTable.AddVariable(Identifier)
        End Sub

        Public Sub AddSymbolInitialization(ByVal symbolTable As SymbolTable)
            symbolTable.AddVariableInitialization(Identifier)
        End Sub

        Public Overrides Sub EmitIL(ByVal scope As CodeGenScope)
            Dim field = scope.Fields(Identifier.NormalizedText)
            scope.ILGenerator.Emit(OpCodes.Ldsfld, field)
        End Sub

        Public Overrides Function ToString() As String
            Return Identifier.Text
        End Function
    End Class
End Namespace
