Imports System.Reflection
Imports System.Reflection.Emit
Imports Microsoft.SmallVisualBasic.Completion
Imports Microsoft.SmallVisualBasic.Expressions
Imports Microsoft.SmallVisualBasic.Library

Namespace Microsoft.SmallVisualBasic.Statements
    Public Class AssignmentStatement
        Inherits Statement

        Public LeftValue As Expression
        Public RightValue As Expression
        Public EqualsToken As Token
        Public IsLocalVariable As Boolean

        Public Overrides Function GetStatementAt(lineNumber As Integer) As Statement
            If lineNumber < StartToken.Line Then Return Nothing
            If RightValue Is Nothing OrElse lineNumber <= RightValue.EndToken.Line Then Return Me
            Return Nothing
        End Function

        Public Overrides Sub AddSymbols(symbolTable As SymbolTable)
            MyBase.AddSymbols(symbolTable)
            EqualsToken.Parent = Me
            Dim isEventHandler = False
            If LeftValue IsNot Nothing Then
                LeftValue.Parent = Me
                Dim prop = TryCast(LeftValue, PropertyExpression)
                If prop IsNot Nothing Then
                    prop.isSet = True
                    If prop.IsDynamic Then
                        Dim typeName = prop.TypeName
                        typeName.Parent = Me ' Important to get the parent subroutine
                        Dim key = symbolTable.GetKey(typeName)
                        symbolTable.InferedTypes(key) = VariableType.Array
                        InferType(symbolTable, prop.PropertyName, prop.DynamicKey)
                    Else
                        Dim typeInfo = symbolTable.GetTypeInfo(prop.TypeName)
                        isEventHandler = typeInfo IsNot Nothing AndAlso
                                typeInfo.Events.TryGetValue(
                                        prop.PropertyName.LCaseText,
                                        Nothing
                                )

                        If Not isEventHandler AndAlso typeInfo IsNot Nothing Then
                            isEventHandler = WinForms.PreCompiler.ContainsEvent(
                                typeInfo.Name,
                                prop.PropertyName.Text
                            )
                        End If

                        prop.IsEvent = isEventHandler
                    End If
                End If

                ' don't move this line up. We must process dynamic property first
                LeftValue.AddSymbols(symbolTable)
            End If

            If RightValue IsNot Nothing Then
                RightValue.Parent = Me
                RightValue.AddSymbols(symbolTable)
                If isEventHandler AndAlso TypeOf RightValue Is IdentifierExpression Then
                    Dim id = CType(RightValue, IdentifierExpression).Identifier
                    symbolTable.PossibleEventHandlers.Add((id, symbolTable.AllIdentifiers.Count - 1))
                End If
            End If

            If EqualsToken.IsIllegal Then Return

            If TypeOf LeftValue Is IdentifierExpression Then
                Dim identifierExp = TryCast(LeftValue, IdentifierExpression)
                Dim key = symbolTable.AddVariable(identifierExp, Me.GetSummery(), IsLocalVariable)
                identifierExp.AddSymbolInitialization(symbolTable)
                InferType(symbolTable, identifierExp.Identifier, key)

            ElseIf TypeOf LeftValue Is ArrayExpression Then
                Dim arrayExpression = CType(LeftValue, ArrayExpression)
                Dim identifier = arrayExpression.LeftHand
                Do While TypeOf identifier Is ArrayExpression
                    identifier = CType(identifier, ArrayExpression).LeftHand
                Loop
                Dim key = symbolTable.AddVariable(identifier, Me.GetSummery(), IsLocalVariable)
                If key <> "" Then
                    symbolTable.InferedTypes(key) = VariableType.Array
                End If
                arrayExpression.AddSymbolInitialization(symbolTable)
            End If
        End Sub

        Private Overloads Sub InferType(
                        symbolTable As SymbolTable,
                        identifier As Token,
                        key As String
                    )

            If key = "" Then Return

            Dim InferedTypes = symbolTable.InferedTypes
            If InferedTypes.ContainsKey(key) Then Return

            Dim varType = WinForms.PreCompiler.GetVarType(identifier.Text)

            Dim type As VariableType
            If varType = VariableType.Any Then
                type = If(RightValue?.InferType(symbolTable), VariableType.Any)
            Else
                type = varType
            End If

            If type <> VariableType.Any Then InferedTypes(key) = type
        End Sub

        Public Overrides Sub EmitIL(scope As CodeGenScope)
            If LeftValue Is Nothing Then Return

            Dim identifierExpr = TryCast(LeftValue, IdentifierExpression)
            Dim propertyExpr = TryCast(LeftValue, PropertyExpression)
            Dim arrayExpr = TryCast(LeftValue, ArrayExpression)
            Dim eventInfo As EventInfo = Nothing

            If identifierExpr IsNot Nothing Then
                If Not LowerAndEmitIL(scope) Then
                    Dim var = scope.GetLocalBuilder(identifierExpr.Subroutine, identifierExpr.Identifier)
                    If var IsNot Nothing Then
                        scope.ILGenerator.Emit(OpCodes.Stloc, var)
                    Else
                        Dim field = scope.Fields(identifierExpr.Identifier.LCaseText)
                        scope.ILGenerator.Emit(OpCodes.Stsfld, field)
                    End If
                End If

            ElseIf arrayExpr IsNot Nothing Then
                If Not LowerAndEmitIL(scope) Then
                    arrayExpr.EmitILForSetter(scope)
                End If

            ElseIf Not scope.ForGlobalHelp AndAlso propertyExpr IsNot Nothing Then
                If propertyExpr.IsDynamic Then
                    Dim code = $"{propertyExpr.TypeName.Text}[""{propertyExpr.PropertyName.Text}""] = {RightValue}"
                    Dim subroutine = SubroutineStatement.GetSubroutine(LeftValue)
                    If subroutine Is Nothing Then subroutine = SubroutineStatement.Current
                    ArrayExpression.ParseAndEmit(code, subroutine, scope, StartToken.Line)
                Else
                    Dim typeInfo = scope.TypeInfoBag.Types(propertyExpr.TypeName.LCaseText)

                    If typeInfo.Events.TryGetValue(propertyExpr.PropertyName.LCaseText, eventInfo) Then
                        Dim identifierExpression2 = TryCast(RightValue, IdentifierExpression)
                        Dim token = scope.SymbolTable.Subroutines(identifierExpression2.Identifier.LCaseText)
                        Dim meth = scope.MethodBuilders(token.LCaseText)
                        scope.ILGenerator.Emit(OpCodes.Ldnull)
                        scope.ILGenerator.Emit(OpCodes.Ldftn, meth)
                        scope.ILGenerator.Emit(OpCodes.Newobj, GetType(SmallVisualBasicCallback).GetConstructors()(0))
                        scope.ILGenerator.EmitCall(OpCodes.Call, eventInfo.GetAddMethod(), Nothing)

                    Else
                        Dim propertyInfo = typeInfo.Properties(propertyExpr.PropertyName.LCaseText)
                        Dim setMethod = propertyInfo.GetSetMethod()
                        RightValue.EmitIL(scope)
                        scope.ILGenerator.EmitCall(OpCodes.Call, setMethod, Nothing)
                    End If

                End If
            End If
        End Sub

        Private Function LowerAndEmitIL(scope As CodeGenScope) As Boolean
            If scope.ForGlobalHelp Then
                ' no need to emit the actaul code. This is just for help info
                LiteralExpression.Zero.EmitIL(scope)
            ElseIf TypeOf RightValue Is InitializerExpression Then
                Dim initExpr = CType(RightValue, InitializerExpression)
                initExpr.LowerAndEmit(LeftValue.ToString(), scope, StartToken.Line)
                Return True
            Else
                RightValue?.EmitIL(scope)
                Return False
            End If
        End Function

        Public Overrides Sub PopulateCompletionItems(
                         bag As CompletionBag,
                         line As Integer,
                         column As Integer,
                         globalScope As Boolean)

            If EqualsToken.IsBefore(line, column) Then
                Dim prop = TryCast(LeftValue, PropertyExpression)
                Dim isHandler = prop IsNot Nothing AndAlso prop.IsEvent
                bag.IsHandler = isHandler
                CompletionHelper.FillSubroutines(bag, functionsOnly:=Not isHandler)
                If Not isHandler Then
                    CompletionHelper.FillExpressionItems(bag)
                End If

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
            Return $"{LeftValue} = {RightValue}" & vbCrLf
        End Function

        Public Overrides Sub InferType(symbolTable As SymbolTable)
            If TypeOf LeftValue Is IdentifierExpression Then
                Dim identifier = TryCast(LeftValue, IdentifierExpression).Identifier
                Dim key = symbolTable.GetKey(identifier)
                InferType(symbolTable, identifier, key)

            ElseIf TypeOf LeftValue Is PropertyExpression Then
                    Dim prop = CType(LeftValue, PropertyExpression)
                If prop.IsDynamic Then
                    InferType(symbolTable, prop.PropertyName, prop.DynamicKey)
                End If
            End If
        End Sub

    End Class
End Namespace
