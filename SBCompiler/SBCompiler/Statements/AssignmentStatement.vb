Imports System.Reflection
Imports System.Reflection.Emit
Imports Microsoft.SmallVisualBasic.Completion
Imports Microsoft.SmallVisualBasic.Engine
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
                    symbolTable.PossibleEventHandlers.Add(New EventHandlerInfo(id, symbolTable.AllIdentifiers.Count - 1))
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

            Dim idExpr = TryCast(LeftValue, IdentifierExpression)

            If idExpr IsNot Nothing Then
                If Not LowerAndEmitIL(scope) Then
                    Dim var = scope.GetLocalBuilder(idExpr.Subroutine, idExpr.Identifier)
                    If var IsNot Nothing Then
                        scope.ILGenerator.Emit(OpCodes.Stloc, var)
                    Else
                        Dim field = scope.Fields(idExpr.Identifier.LCaseText)
                        scope.ILGenerator.Emit(OpCodes.Stsfld, field)
                    End If
                End If
                Exit Sub
            End If

            Dim arrExpr = TryCast(LeftValue, ArrayExpression)
            If arrExpr IsNot Nothing Then
                If Not LowerAndEmitIL(scope) Then
                    arrExpr.EmitILForSetter(scope)
                End If
                Exit Sub
            End If

            Dim propExpr = TryCast(LeftValue, PropertyExpression)
            If Not scope.ForGlobalHelp AndAlso propExpr IsNot Nothing Then
                If propExpr.IsDynamic Then
                    Dim subroutine = SubroutineStatement.GetSubroutine(LeftValue)
                    Dim code = $"{propExpr.TypeName.Text}[""{propExpr.PropertyName.Text}""] = {RightValue}"
                    If subroutine Is Nothing Then subroutine = SubroutineStatement.Current
                    Parser.ParseAndEmit(code, subroutine, scope, StartToken.Line)
                Else
                    Dim typeInfo = scope.TypeInfoBag.Types(propExpr.TypeName.LCaseText)
                    Dim eventInfo As EventInfo = Nothing

                    If typeInfo.Events.TryGetValue(propExpr.PropertyName.LCaseText, EventInfo) Then
                        If RightValue.StartToken.Type = TokenType.Nothing Then
                            scope.ILGenerator.Emit(OpCodes.Ldnull)
                            scope.ILGenerator.Emit(OpCodes.Ldftn, GetType(SmallVisualBasic.Library.Program).GetMethod("DoNothing", BindingFlags.Public Or BindingFlags.Static))
                            scope.ILGenerator.Emit(OpCodes.Newobj, GetType(SmallVisualBasicCallback).GetConstructors()(0))
                            scope.ILGenerator.EmitCall(OpCodes.Call, EventInfo.GetRemoveMethod(), Nothing)
                        Else
                            Dim subExpr = TryCast(RightValue, IdentifierExpression)
                            Dim subName = scope.SymbolTable.Subroutines(subExpr.Identifier.LCaseText)
                            Dim method = scope.MethodBuilders(subName.LCaseText)
                            scope.ILGenerator.Emit(OpCodes.Ldnull)
                            scope.ILGenerator.Emit(OpCodes.Ldftn, method)
                            scope.ILGenerator.Emit(OpCodes.Newobj, GetType(SmallVisualBasicCallback).GetConstructors()(0))
                            scope.ILGenerator.EmitCall(OpCodes.Call, EventInfo.GetAddMethod(), Nothing)
                        End If

                    Else
                        Dim propertyInfo = typeInfo.Properties(propExpr.PropertyName.LCaseText)
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
                If isHandler Then
                    bag.CompletionItems.Add(New CompletionItem() With {
                        .DisplayName = "Nothing",
                        .ItemType = CompletionItemType.Keyword,
                        .Key = "nothing",
                        .ReplacementText = "Nothing"
                    })
                Else
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

        Public Overrides Function Execute(runner As ProgramRunner) As Statement
            Dim idExpr = TryCast(LeftValue, IdentifierExpression)
            If idExpr IsNot Nothing Then
                runner.Fields(runner.GetKey(idExpr.Identifier)) = RightValue.Evaluate(runner)
                Return Nothing
            End If

            Dim arrExpr = TryCast(LeftValue, ArrayExpression)
            If arrExpr IsNot Nothing Then
                SetArrayValue(runner, arrExpr, RightValue.Evaluate(runner))
                Return Nothing
            End If

            Dim propExpr = TryCast(LeftValue, PropertyExpression)
            If propExpr IsNot Nothing Then
                If propExpr.IsDynamic Then
                    Dim arrExpr2 As New ArrayExpression() With {
                        .LeftHand = New IdentifierExpression() With {.Identifier = propExpr.TypeName},
                        .Indexer = New IdentifierExpression() With {.Identifier = propExpr.PropertyName},
                        .Parent = Me
                    }
                    SetArrayValue(runner, arrExpr2, RightValue.Evaluate(runner))

                Else
                    Dim typeInfo = runner.TypeInfoBag.Types(propExpr.TypeName.LCaseText)
                    Dim eventInfo As EventInfo = Nothing

                    If typeInfo.Events.TryGetValue(propExpr.PropertyName.LCaseText, eventInfo) Then
                        If RightValue.StartToken.Type = TokenType.Nothing Then
                            eventInfo.GetAddMethod().Invoke(Nothing, {CType(Sub() Exit Sub, SmallVisualBasicCallback)})
                        Else
                            Dim subExpr = TryCast(RightValue, IdentifierExpression)
                            Dim subroutine = runner.SymbolTable.Subroutines(subExpr.Identifier.LCaseText)
                            eventInfo.GetAddMethod().Invoke(Nothing,
                                     {CType(Sub() subroutine.Parent.Execute(runner), SmallVisualBasicCallback)}
                            )
                        End If

                    Else
                        Dim propertyInfo = typeInfo.Properties(propExpr.PropertyName.LCaseText)
                        propertyInfo.SetValue(Nothing, RightValue.Evaluate(runner), Nothing)
                    End If
                End If
            End If

            Return Nothing
        End Function


        Friend Sub SetArrayValue(runner As ProgramRunner, arrExpr As ArrayExpression, value As Primitive)
            Dim idExpr = TryCast(arrExpr.LeftHand, IdentifierExpression)
            Dim fields = runner.Fields

            If idExpr IsNot Nothing Then
                Dim arrName = runner.GetKey(idExpr.Identifier)
                Dim arr As Primitive = Nothing
                If Not fields.TryGetValue(arrName, arr) Then
                    arr = Nothing
                End If
                fields(arrName) = Primitive.SetArrayValue(value, arr, arrExpr.Indexer.Evaluate(runner))

            Else
                Dim arrExpr2 = TryCast(arrExpr.LeftHand, ArrayExpression)
                If arrExpr2 IsNot Nothing Then
                    Dim arr2 = Primitive.SetArrayValue(
                          value,
                          arrExpr.Evaluate(runner),
                          arrExpr.Indexer.Evaluate(runner)
                    )
                    SetArrayValue(runner, arrExpr2, arr2)
                End If
            End If
        End Sub

    End Class
End Namespace
