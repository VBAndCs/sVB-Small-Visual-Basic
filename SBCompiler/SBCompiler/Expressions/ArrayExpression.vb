Imports System
Imports System.Globalization
Imports System.Reflection.Emit

Namespace Microsoft.SmallBasic.Expressions
    <Serializable>
    Public Class ArrayExpression
        Inherits Expression

        Public Property LeftHand As Expression
        Public Property Indexer As Expression

        Public Overrides Sub AddSymbols(symbolTable As SymbolTable)
            MyBase.AddSymbols(symbolTable)
            If LeftHand IsNot Nothing Then
                LeftHand.Parent = Me.Parent
                LeftHand.AddSymbols(symbolTable)
            End If

            If Indexer IsNot Nothing Then
                Indexer.Parent = Me.Parent
                Indexer.AddSymbols(symbolTable)
            End If
        End Sub

        Public Sub AddSymbolInitialization(symbolTable As SymbolTable)
            Dim arrayExpression = TryCast(LeftHand, ArrayExpression)
            Dim identifierExpression = TryCast(LeftHand, IdentifierExpression)

            If arrayExpression IsNot Nothing Then
                arrayExpression.AddSymbolInitialization(symbolTable)
            Else
                identifierExpression?.AddSymbolInitialization(symbolTable)
            End If
        End Sub

        Public Overrides Sub EmitIL(scope As CodeGenScope)
            LeftHand.EmitIL(scope)
            Indexer.EmitIL(scope)
            scope.ILGenerator.EmitCall(OpCodes.Call, scope.TypeInfoBag.GetArrayValue, Nothing)
        End Sub

        Public Sub EmitILForSetter(scope As CodeGenScope)
            LeftHand.EmitIL(scope)
            Indexer.EmitIL(scope)
            scope.ILGenerator.EmitCall(OpCodes.Call, scope.TypeInfoBag.SetArrayValue, Nothing)
            Dim arrayExpression = TryCast(LeftHand, ArrayExpression)
            Dim identifierExpression = TryCast(LeftHand, IdentifierExpression)

            If arrayExpression IsNot Nothing Then
                arrayExpression.EmitILForSetter(scope)
            ElseIf identifierExpression IsNot Nothing Then
                Dim subroutine = identifierExpression.Subroutine
                Dim identifier = identifierExpression.Identifier

                Dim var = scope.GetLocalBuilder(subroutine, identifier)
                If var IsNot Nothing Then
                    scope.ILGenerator.Emit(OpCodes.Stloc, var)

                ElseIf scope.Fields.ContainsKey(identifier.NormalizedText) Then
                    Dim field = scope.Fields(identifierExpression.Identifier.NormalizedText)
                    scope.ILGenerator.Emit(OpCodes.Stsfld, field)

                Else 'If Not CodeGenerator.IgnoreVarErrors Then
                    scope.SymbolTable.Errors.Add(New [Error](identifier, $"The variable `{identifier.Text}` is used before beeing initialized."))
                End If

            End If
        End Sub

        Public Overrides Function ToString() As String
            Return String.Format(CultureInfo.CurrentUICulture, "{0}[{1}]", New Object(1) {LeftHand.ToString(), Indexer})
        End Function
    End Class
End Namespace
