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
        Public IsLocalVariable As Boolean

        Public Overrides Sub AddSymbols(symbolTable As SymbolTable)
            MyBase.AddSymbols(symbolTable)
            EqualsToken.Parent = Me

            If LeftValue IsNot Nothing Then
                LeftValue.Parent = Me
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
        End Sub

        Private Function LowerAndEmitIL(scope As CodeGenScope) As Boolean
            If TypeOf RightValue Is InitializerExpression Then
                Dim initExpr = CType(RightValue, InitializerExpression)
                Dim code = initExpr.Lower(LeftValue.ToString())

                Dim _parser = Parser.Parse(code)
                Dim semantic As New SemanticAnalyzer(_parser, scope.TypeInfoBag)
                semantic.Analyze()
                Dim _newSymbolTable = _parser.SymbolTable
                Dim _oldSymbolTable = scope.SymbolTable
                _oldSymbolTable.CopyFrom(_newSymbolTable)

                'Build new fields
                Dim symbolTable = _parser.SymbolTable
                For Each key In symbolTable.Variables.Keys
                    Dim value As FieldInfo = scope.TypeBuilder.DefineField(key, GetType(Primitive), FieldAttributes.Private Or FieldAttributes.Static)
                    If Not scope.Fields.ContainsKey(key) Then scope.Fields.Add(key, value)
                Next

                ' EmitIL
                For Each item In _parser.ParseTree
                    item.PrepareForEmit(scope)
                Next

                For Each item In _parser.ParseTree
                    item.EmitIL(scope)
                Next
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
