Imports System.Globalization
Imports System.Reflection
Imports System.Reflection.Emit
Imports Microsoft.SmallBasic.Completion
Imports Microsoft.SmallBasic.Expressions
Imports Microsoft.SmallBasic.Library

Namespace Microsoft.SmallBasic.Statements
    Public Class AssignmentStatement
        Inherits Statement

        Public LeftValue As Expression
        Public RightValue As Expression
        Public EqualsToken As TokenInfo

        Public Overrides Sub AddSymbols(ByVal symbolTable As SymbolTable)
            If LeftValue IsNot Nothing Then
                LeftValue.AddSymbols(symbolTable)
            End If

            If RightValue IsNot Nothing Then
                RightValue.AddSymbols(symbolTable)
            End If

            Dim identifierExpression As IdentifierExpression = TryCast(LeftValue, IdentifierExpression)
            Dim arrayExpression As ArrayExpression = TryCast(LeftValue, ArrayExpression)

            If identifierExpression IsNot Nothing Then
                identifierExpression.AddSymbolInitialization(symbolTable)
            Else
                arrayExpression?.AddSymbolInitialization(symbolTable)
            End If
        End Sub

        Public Overrides Sub EmitIL(ByVal scope As CodeGenScope)
            Dim identifierExpression As IdentifierExpression = TryCast(LeftValue, IdentifierExpression)
            Dim propertyExpression As PropertyExpression = TryCast(LeftValue, PropertyExpression)
            Dim arrayExpression As ArrayExpression = TryCast(LeftValue, ArrayExpression)
            Dim value As EventInfo = Nothing

            If identifierExpression IsNot Nothing Then
                RightValue.EmitIL(scope)
                Dim field = scope.Fields(identifierExpression.Identifier.NormalizedText)
                scope.ILGenerator.Emit(OpCodes.Stsfld, field)
            ElseIf arrayExpression IsNot Nothing Then
                RightValue.EmitIL(scope)
                arrayExpression.EmitILForSetter(scope)
            ElseIf propertyExpression IsNot Nothing Then
                Dim typeInfo = scope.TypeInfoBag.Types(propertyExpression.TypeName.NormalizedText)

                If typeInfo.Events.TryGetValue(propertyExpression.PropertyName.NormalizedText, value) Then
                    Dim identifierExpression2 As IdentifierExpression = TryCast(RightValue, IdentifierExpression)
                    Dim tokenInfo = scope.SymbolTable.Subroutines(identifierExpression2.Identifier.NormalizedText)
                    Dim meth = scope.MethodBuilders(tokenInfo.NormalizedText)
                    scope.ILGenerator.Emit(OpCodes.Ldnull)
                    scope.ILGenerator.Emit(OpCodes.Ldftn, meth)
                    scope.ILGenerator.Emit(OpCodes.Newobj, GetType(SmallBasicCallback).GetConstructors()(0))
                    scope.ILGenerator.EmitCall(OpCodes.Call, value.GetAddMethod(), Nothing)
                Else
                    Dim propertyInfo = typeInfo.Properties(propertyExpression.PropertyName.NormalizedText)
                    Dim setMethod As MethodInfo = propertyInfo.GetSetMethod()
                    RightValue.EmitIL(scope)
                    scope.ILGenerator.EmitCall(OpCodes.Call, setMethod, Nothing)
                End If
            End If
        End Sub

        Public Overrides Sub PopulateCompletionItems(ByVal completionBag As CompletionBag, ByVal line As Integer, ByVal column As Integer, ByVal globalScope As Boolean)
            If EqualsToken.Token = Token.Illegal OrElse column <= EqualsToken.Column Then
                CompletionHelper.FillAllGlobalItems(completionBag, globalScope)
            Else
                CompletionHelper.FillExpressionItems(completionBag)
            End If
        End Sub

        Public Overrides Function ToString() As String
            Return String.Format(CultureInfo.CurrentUICulture, "{0} {1} {2}" & VisualBasic.Constants.vbCrLf, New Object(2) {LeftValue, EqualsToken.Text, RightValue})
        End Function
    End Class
End Namespace
