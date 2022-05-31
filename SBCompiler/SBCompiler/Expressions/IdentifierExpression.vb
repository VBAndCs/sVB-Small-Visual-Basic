Imports System
Imports System.Reflection.Emit
Imports Microsoft.SmallBasic.Statements

Namespace Microsoft.SmallBasic.Expressions
    <Serializable>
    Public Class IdentifierExpression
        Inherits Expression

        Public Property Identifier As TokenInfo
        Public Subroutine As SubroutineStatement

        Public Overrides Sub AddSymbols(symbolTable As SymbolTable)
            MyBase.AddSymbols(symbolTable)
            _Identifier.Parent = Me.Parent
        End Sub

        Public Sub AddSymbolInitialization(ByVal symbolTable As SymbolTable)
            symbolTable.AddVariableInitialization(Identifier)
        End Sub

        Public Overrides Sub EmitIL(ByVal scope As CodeGenScope)
            Dim var = scope.GetLocalBuilder(Subroutine, Identifier)

            If var IsNot Nothing Then
                scope.ILGenerator.Emit(OpCodes.Ldloc, var)

            ElseIf scope.Fields.ContainsKey(Identifier.NormalizedText) Then
                Dim field = scope.Fields(Identifier.NormalizedText)
                scope.ILGenerator.Emit(OpCodes.Ldsfld, field)

            ElseIf Not CodeGenerator.IgnoreVarErrors Then
                scope.SymbolTable.Errors.Add(New [Error](Identifier, $"The variable `{Identifier.Text}` is used before being initialized."))
            End If
        End Sub

        Public Overrides Function ToString() As String
            Return Identifier.Text
        End Function
    End Class
End Namespace
