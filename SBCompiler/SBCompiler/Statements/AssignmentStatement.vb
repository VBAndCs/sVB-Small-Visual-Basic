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
        Public EqualsToken As Token
        Public IsLocalVariable As Boolean

        Public Overrides Function GetStatementAt(lineNumber As Integer) As Statement
            If lineNumber < StartToken.Line Then Return Nothing
            If lineNumber <= RightValue?.EndToken.Line Then Return Me
            Return Nothing
        End Function

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
                If TypeOf RightValue Is IdentifierExpression Then
                    Dim id = CType(RightValue, IdentifierExpression).Identifier
                    symbolTable.PossibleEventHandlers.Add((id, symbolTable.AllIdentifiers.Count - 1))
                End If
            End If

            If TypeOf LeftValue Is IdentifierExpression Then
                Dim identifierExpression = TryCast(LeftValue, IdentifierExpression)
                Dim key = symbolTable.AddVariable(identifierExpression, Me.GetSummery(), IsLocalVariable)
                If key <> "" Then
                    Dim varType = WinForms.PreCompiler.GetVarType(identifierExpression.Identifier.Text)
                    If varType <> VariableType.None Then
                        symbolTable.InferedTypes(key) = varType
                    Else
                        Dim literalExpr = TryCast(RightValue, LiteralExpression)
                        If literalExpr IsNot Nothing Then
                            Select Case literalExpr.Literal.Type
                                Case TokenType.NumericLiteral
                                    symbolTable.InferedTypes(key) = VariableType.Double
                                Case TokenType.True, TokenType.False
                                    symbolTable.InferedTypes(key) = VariableType.Boolean
                                Case TokenType.StringLiteral
                                    symbolTable.InferedTypes(key) = VariableType.String
                                Case Else
                                    Dim name = identifierExpression.Identifier.NormalizedText
                                    If name.StartsWith("color") OrElse name.EndsWith("color") OrElse name.EndsWith("colors") Then
                                        symbolTable.InferedTypes(key) = VariableType.Color
                                    End If
                            End Select

                        ElseIf TypeOf RightValue Is NegativeExpression Then
                            symbolTable.InferedTypes(key) = VariableType.Double

                        ElseIf TypeOf RightValue Is InitializerExpression Then
                            symbolTable.InferedTypes(key) = VariableType.Array

                        ElseIf TypeOf RightValue Is IdentifierExpression Then
                            symbolTable.InferedTypes(key) = symbolTable.GetInferedType(CType(RightValue, IdentifierExpression).Identifier)

                        ElseIf TypeOf RightValue Is MethodCallExpression Then
                            Dim methodExpr = CType(RightValue, MethodCallExpression)
                            symbolTable.InferedTypes(key) = symbolTable.GetReturnValueType(methodExpr.TypeName, methodExpr.MethodName, True)

                        ElseIf TypeOf RightValue Is PropertyExpression Then
                            Dim propExpr = CType(RightValue, PropertyExpression)
                            symbolTable.InferedTypes(key) = symbolTable.GetReturnValueType(propExpr.TypeName, propExpr.PropertyName, False)
                        Else
                            Dim name = identifierExpression.Identifier.NormalizedText
                            If name.StartsWith("color") OrElse name.EndsWith("color") OrElse name.EndsWith("colors") Then
                                symbolTable.InferedTypes(key) = VariableType.Color
                            End If
                        End If
                    End If
                End If

                identifierExpression.AddSymbolInitialization(symbolTable)

            ElseIf TypeOf LeftValue Is ArrayExpression Then
                Dim arrayExpression = CType(LeftValue, ArrayExpression)
                Dim key = symbolTable.AddVariable(arrayExpression.LeftHand, Me.GetSummery(), IsLocalVariable)
                If TypeOf arrayExpression.LeftHand Is IdentifierExpression Then
                    If key <> "" Then
                        symbolTable.InferedTypes(key) = VariableType.Array
                    End If
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
                    ArrayExpression.ParseAndEmit(code, subroutine, scope, StartToken.Line)
                Else
                    Dim typeInfo = scope.TypeInfoBag.Types(propertyExpression.TypeName.NormalizedText)

                    If typeInfo.Events.TryGetValue(propertyExpression.PropertyName.NormalizedText, value) Then
                        Dim identifierExpression2 = TryCast(RightValue, IdentifierExpression)
                        Dim token = scope.SymbolTable.Subroutines(identifierExpression2.Identifier.NormalizedText)
                        Dim meth = scope.MethodBuilders(token.NormalizedText)
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
                initExpr.LowerAndEmit(LeftValue.ToString(), scope, StartToken.Line)
                Return True
            Else
                RightValue.EmitIL(scope)
                Return False
            End If
        End Function

        Public Overrides Sub PopulateCompletionItems(
                         bag As CompletionBag,
                         line As Integer,
                         column As Integer,
                         globalScope As Boolean)

            If EqualsToken.IsBefore(line, column) Then
                CompletionHelper.FillExpressionItems(bag)
                Dim prop = TryCast(LeftValue, PropertyExpression)
                Dim canBeAHandler = bag.NextToEquals AndAlso prop IsNot Nothing AndAlso Not prop.IsDynamic
                CompletionHelper.FillSubroutines(bag, functionsOnly:=Not canBeAHandler)
            Else
                Dim arrayExpr = TryCast(LeftValue, ArrayExpression)
                If arrayExpr Is Nothing Then
                    bag.IsFirstToken = StartToken.Contains(line, column, True)
                    CompletionHelper.FillAllGlobalItems(bag, globalScope)
                    bag.IsFirstToken = False
                Else
                    Dim indexer = arrayExpr.Indexer
                    If indexer IsNot Nothing AndAlso indexer.StartToken.IsBefore(line, column) Then
                        CompletionHelper.FillAllGlobalItems(bag, globalScope)
                    Else
                        CompletionHelper.FillExpressionItems(bag)
                        CompletionHelper.FillSubroutines(bag, functionsOnly:=True)
                    End If
                End If
            End If
        End Sub

        Public Overrides Function ToString() As String
            Return String.Format(CultureInfo.CurrentUICulture, "{0} {1} {2}" & VisualBasic.Constants.vbCrLf, New Object(2) {LeftValue, EqualsToken.Text, RightValue})
        End Function


    End Class
End Namespace
