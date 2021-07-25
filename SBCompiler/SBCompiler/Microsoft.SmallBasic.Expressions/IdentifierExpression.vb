Imports System
Imports System.Reflection.Emit
Imports Microsoft.SmallBasic.Statements

Namespace Microsoft.SmallBasic.Expressions
    <Serializable>
    Public Class IdentifierExpression
        Inherits Expression

        Public Property Identifier As TokenInfo
        Public Subroutine As SubroutineStatement


        Public Sub AddSymbolInitialization(ByVal symbolTable As SymbolTable)
            symbolTable.AddVariableInitialization(Identifier)
        End Sub

        Public Overrides Sub EmitIL(ByVal scope As CodeGenScope)
            If scope.Fields.ContainsKey(Identifier.NormalizedText) Then
                Dim field = scope.Fields(Identifier.NormalizedText)
                scope.ILGenerator.Emit(OpCodes.Ldsfld, field)
            Else
                Dim localkey = ""
                If Subroutine Is Nothing Then
                    localkey = Identifier.NormalizedText
                Else
                    localkey = $"{Subroutine.Name.NormalizedText}.{Identifier.NormalizedText}"
                End If

                If scope.Locals.ContainsKey(localkey) Then
                    Dim var = scope.Locals(localkey)
                    scope.ILGenerator.Emit(OpCodes.Ldloc, var)
                Else
                    scope.SymbolTable.Errors.Add(New [Error](Identifier, $"The variable `{Identifier.Text}` is used but before beeing initialized."))
                End If
            End If
        End Sub

        Public Overrides Function ToString() As String
            Return Identifier.Text
        End Function
    End Class
End Namespace
