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

            Dim varName = identifier.Text
            Dim varType = WinForms.PreCompiler.GetVarType(varName)
            Dim varType2 = If(RightValue?.InferType(symbolTable), VariableType.Any)
            Dim type = If(varType = VariableType.Any OrElse
                        (varType >= VariableType.Control AndAlso varName.Length < varType.ToString().Length AndAlso varType2 < VariableType.Control),
                  varType2,
                  varType
            )
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
                            scope.ILGenerator.EmitCall(OpCodes.Call, eventInfo.GetAddMethod(), Nothing)
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
                Return False
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
                Dim value = RightValue.Evaluate(runner)
                If propExpr.IsDynamic Then
                    Dim arrExpr2 As New ArrayExpression() With {
                        .LeftHand = New IdentifierExpression() With {.Identifier = propExpr.TypeName},
                        .Indexer = New LiteralExpression($"""{propExpr.PropertyName.Text}"""),
                        .Parent = Me
                    }
                    SetArrayValue(runner, arrExpr2, value)

                ElseIf propExpr.TypeName.LCaseText = "global" Then
                    runner.SetGlobalField(propExpr.PropertyName.LCaseText, value)

                Else
                    Dim typeInfo = runner.TypeInfoBag.Types(propExpr.TypeName.LCaseText)
                    Dim eventInfo As EventInfo = Nothing

                    If typeInfo.Events.TryGetValue(propExpr.PropertyName.LCaseText, eventInfo) Then
                        If RightValue.StartToken.Type = TokenType.Nothing Then
                            eventInfo.GetRemoveMethod().Invoke(Nothing, {CType(Sub() Exit Sub, SmallVisualBasicCallback)})
                        Else
                            SetEventHandler(runner, typeInfo, eventInfo)
                        End If

                    Else
                        Dim propertyInfo = typeInfo.Properties(propExpr.PropertyName.LCaseText)
                        propertyInfo.SetValue(Nothing, value, Nothing)
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
                fields.TryGetValue(arrName, arr)
                fields(arrName) = Primitive.SetArrayValue(value, arr, arrExpr.Indexer.Evaluate(runner))

            Else
                Dim arrExpr2 = TryCast(arrExpr.LeftHand, ArrayExpression)
                If arrExpr2 IsNot Nothing Then
                    Dim arr2 = Primitive.SetArrayValue(
                          value,
                          arrExpr2.Evaluate(runner),
                          arrExpr.Indexer.Evaluate(runner)
                    )
                    SetArrayValue(runner, arrExpr2, arr2)
                End If
            End If
        End Sub

        Private Sub SetEventHandler(runner As ProgramRunner, typeInfo As TypeInfo, eventInfo As EventInfo)
            Dim subExpr = TryCast(RightValue, IdentifierExpression)
            Dim subroutine = runner.SymbolTable.Subroutines(subExpr.Identifier.LCaseText)
            Dim handlerKey As String
            Dim engine = runner.Engine

            If typeInfo.Name = "Thread" AndAlso eventInfo.Name = "SubToRun" Then
                subroutine.Parent.Execute(runner)
                Return
            End If

            If typeInfo.Type.FullName.StartsWith("Microsoft.SmallVisualBasic.WinForms") Then
                handlerKey = $"{WinForms.Event.SenderControl}.{eventInfo.Name}".ToLower()
            Else
                handlerKey = $"{typeInfo.Key}.{eventInfo.Name}".ToLower()
            End If

            Dim handlerRunner = runner.CreateSubRunner()
            eventInfo.GetAddMethod().Invoke(Nothing, {CType(
                   Sub()
                       Dim eventName = eventInfo.Name
                       If engine.BreakMode AndAlso (eventName.EndsWith("Click") OrElse eventName.Contains("Key") OrElse eventName.Contains("Mouse")) Then
                           MsgBox("You are running the program in debug mode, and it is curreuntly paused, so you can't do anything before resuming execution. You will be switched to the current line that has the break point.")
                           engine.RaiseDebuggerStateChangedEvent(Nothing)
                           Return
                       End If

                       If runner.EventThreads.ContainsKey(handlerKey) Then
                           If runner.EventThreads(handlerKey) Then Return
                       End If

                       runner.EventThreads(handlerKey) = True
                       Dim invokerRunner = engine.CurrentRunner
                       Dim dc = invokerRunner.DebuggerCommand
                       Dim isTimerTick = typeInfo.Name.EndsWith("Timer") AndAlso eventInfo.Name.EndsWith("Tick")

                       If Not isTimerTick AndAlso invokerRunner.DebuggerState = DebuggerState.Running Then
                           Try
                               invokerRunner.runnerThread.Suspend()
                               invokerRunner.IsPaused = True
                           Catch ex As Exception
                           End Try
                       End If

                       Dim handlerThread As New Threading.Thread(
                            Sub()
                                handlerRunner.HasBeenPaused = False
                                handlerRunner.StepAround = False
                                handlerRunner.Depth = 0
                                handlerRunner.Breakpoints = runner.Breakpoints
                                handlerRunner.DebuggerState = DebuggerState.Running

                                If dc = DebuggerCommand.Run AndAlso ProgramRunner.EventsCommand = DebuggerCommand.Run Then
                                    handlerRunner.DebuggerCommand = DebuggerCommand.Run
                                Else
                                    handlerRunner.DebuggerCommand = DebuggerCommand.StepInto
                                    ProgramRunner.EventsCommand = DebuggerCommand.Run
                                End If

                                engine.CurrentRunner = handlerRunner
                                subroutine.Parent.Execute(handlerRunner)

                                dc = handlerRunner.DebuggerCommand
                                handlerRunner.DebuggerCommand = DebuggerCommand.Run
                                If dc = DebuggerCommand.Run Then
                                    ProgramRunner.EventsCommand = dc
                                Else
                                    ProgramRunner.EventsCommand = DebuggerCommand.StepInto
                                    invokerRunner.StepAround = False
                                    invokerRunner.Depth = 0
                                End If

                                handlerRunner.StepAround = False
                                handlerRunner.ChangeDebuggerState(DebuggerState.Finished)
                                handlerRunner.ContinueAll()
                                If invokerRunner.IsPaused Then
                                    engine.PausedRunner = invokerRunner
                                    If ProgramRunner.EventsCommand = DebuggerCommand.Run Then
                                        invokerRunner.DebuggerCommand = DebuggerCommand.Run
                                        invokerRunner.Continue()
                                    Else
                                        ProgramRunner.EventsCommand = DebuggerCommand.Run
                                        invokerRunner.DebuggerCommand = DebuggerCommand.StepInto
                                        invokerRunner.ChangeDebuggerState(DebuggerState.Paused)
                                    End If
                                End If
                                runner.EventThreads(handlerKey) = False
                            End Sub)

                       handlerThread.IsBackground = True
                       handlerRunner.runnerThread = handlerThread
                       handlerThread.Start()
                   End Sub, SmallVisualBasicCallback)
            })
        End Sub

    End Class
End Namespace
