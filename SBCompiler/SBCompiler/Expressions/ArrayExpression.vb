﻿Imports System
Imports System.Globalization
Imports System.Reflection
Imports System.Reflection.Emit
Imports Microsoft.SmallBasic.Library
Imports Microsoft.SmallBasic.Statements

Namespace Microsoft.SmallBasic.Expressions
    <Serializable>
    Public Class ArrayExpression
        Inherits Expression

        Public Property LeftHand As Expression
        Public Property Indexer As Expression

        Public Shared Function ParseAndEmit(code As String, subroutine As SubroutineStatement, scope As CodeGenScope) As Expression
            Dim tempRoutine = SubroutineStatement.Current
            SubroutineStatement.Current = subroutine
            Dim _parser = Parser.Parse(code, scope.SymbolTable, scope.TypeInfoBag)

            Dim semantic As New SemanticAnalyzer(_parser, scope.TypeInfoBag)
            semantic.Analyze()

            'Build new fields
            For Each key In _parser.SymbolTable.Variables.Keys
                If Not scope.Fields.ContainsKey(key) Then
                    Dim value = scope.TypeBuilder.DefineField(key, GetType(Primitive), FieldAttributes.Private Or FieldAttributes.Static)
                    scope.Fields.Add(key, value)
                End If
            Next

            ' EmitIL
            For Each item In _parser.ParseTree
                item.PrepareForEmit(scope)
            Next

            For Each item In _parser.ParseTree
                item.EmitIL(scope)
            Next

            Statements.SubroutineStatement.Current = tempRoutine

            Return CType(_parser.ParseTree(0), Statements.AssignmentStatement).LeftValue
        End Function


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

                Else
                    scope.SymbolTable.Errors.Add(New [Error](identifier, $"The variable `{identifier.Text}` is used before being initialized."))
                End If

            End If
        End Sub

        Public Overrides Function ToString() As String
            Return String.Format(CultureInfo.CurrentUICulture, "{0}[{1}]", New Object(1) {LeftHand.ToString(), Indexer})
        End Function
    End Class
End Namespace
