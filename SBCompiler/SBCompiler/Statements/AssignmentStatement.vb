﻿Imports System.Globalization
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
        Public IsLocalVariable As Boolean

        Public Overrides Sub AddSymbols(symbolTable As SymbolTable)
            MyBase.AddSymbols(symbolTable)
            EqualsToken.Parent = Me

            If LeftValue IsNot Nothing Then
                LeftValue.Parent = Me
                Dim prop = TryCast(LeftValue, PropertyExpression)
                If prop IsNot Nothing Then
                    prop.isSet = True
                End If
                LeftValue.AddSymbols(symbolTable)
            End If

            If RightValue IsNot Nothing Then
                RightValue.Parent = Me
                RightValue.AddSymbols(symbolTable)
            End If

            If TypeOf LeftValue Is IdentifierExpression Then
                Dim identifierExpression = CType(LeftValue, IdentifierExpression)
                symbolTable.AddVariable(identifierExpression, IsLocalVariable)
                identifierExpression.AddSymbolInitialization(symbolTable)

            ElseIf TypeOf LeftValue Is ArrayExpression Then
                Dim arrayExpression = CType(LeftValue, ArrayExpression)
                If TypeOf arrayExpression.LeftHand Is IdentifierExpression Then
                    symbolTable.AddVariable(arrayExpression.LeftHand, IsLocalVariable)
                End If
                arrayExpression.AddSymbolInitialization(symbolTable)
                End If
        End Sub

        Public Overrides Sub EmitIL(scope As CodeGenScope)
            Dim identifierExpression = TryCast(LeftValue, IdentifierExpression)
            Dim propertyExpression = TryCast(LeftValue, PropertyExpression)
            Dim arrayExpression = TryCast(LeftValue, ArrayExpression)
            Dim value As EventInfo = Nothing

            If identifierExpression IsNot Nothing Then
                If Not LowerAndEmitIL(scope) Then
                    Dim var = scope.GetLocalBuilder(identifierExpression.Subroutine, identifierExpression.Identifier)
                    If var IsNot Nothing Then
                        scope.ILGenerator.Emit(OpCodes.Stloc, var)
                    Else
                        Dim field = scope.Fields(identifierExpression.Identifier.NormalizedText)
                        scope.ILGenerator.Emit(OpCodes.Stsfld, field)
                    End If
                End If

            ElseIf arrayExpression IsNot Nothing Then
                If Not LowerAndEmitIL(scope) Then
                    arrayExpression.EmitILForSetter(scope)
                End If

            ElseIf propertyExpression IsNot Nothing Then
                If propertyExpression.IsDynamic Then
                    Dim code = $"{propertyExpression.TypeName.Text}[""{propertyExpression.PropertyName.Text}""] = {RightValue}"
                    Dim subroutine = SubroutineStatement.GetSubroutine(LeftValue)
                    If subroutine Is Nothing Then subroutine = SubroutineStatement.Current
                    ArrayExpression.ParseAndEmit(code, subroutine, scope)
                Else
                    Dim typeInfo = scope.TypeInfoBag.Types(propertyExpression.TypeName.NormalizedText)

                    If typeInfo.Events.TryGetValue(propertyExpression.PropertyName.NormalizedText, value) Then
                        Dim identifierExpression2 = TryCast(RightValue, IdentifierExpression)
                        Dim tokenInfo = scope.SymbolTable.Subroutines(identifierExpression2.Identifier.NormalizedText)
                        Dim meth = scope.MethodBuilders(tokenInfo.NormalizedText)
                        scope.ILGenerator.Emit(OpCodes.Ldnull)
                        scope.ILGenerator.Emit(OpCodes.Ldftn, meth)
                        scope.ILGenerator.Emit(OpCodes.Newobj, GetType(SmallBasicCallback).GetConstructors()(0))
                        scope.ILGenerator.EmitCall(OpCodes.Call, value.GetAddMethod(), Nothing)

                    Else
                        Dim propertyInfo = typeInfo.Properties(propertyExpression.PropertyName.NormalizedText)
                        Dim setMethod = propertyInfo.GetSetMethod()
                        RightValue.EmitIL(scope)
                        scope.ILGenerator.EmitCall(OpCodes.Call, setMethod, Nothing)
                    End If

                End If
            End If
        End Sub

        Private Function LowerAndEmitIL(scope As CodeGenScope) As Boolean
            If TypeOf RightValue Is InitializerExpression Then
                Dim initExpr = CType(RightValue, InitializerExpression)
                initExpr.LowerAndEmit(LeftValue.ToString(), scope)
                Return True
            Else
                RightValue.EmitIL(scope)
                Return False
            End If
        End Function

        Public Overrides Sub PopulateCompletionItems(
                         completionBag As CompletionBag,
                         line As Integer,
                         column As Integer,
                         globalScope As Boolean)

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
