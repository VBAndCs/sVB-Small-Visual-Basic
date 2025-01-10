Imports System.Globalization
Imports System.Reflection.Emit
Imports Microsoft.SmallVisualBasic.Completion
Imports Microsoft.SmallVisualBasic.Engine
Imports Microsoft.SmallVisualBasic.Expressions

Namespace Microsoft.SmallVisualBasic.Statements
    Public Class SubroutineCallStatement
        Inherits Statement

        Public Name As Token
        Public Args As List(Of Expression)
        Public OuterSubroutine As SubroutineStatement
        Friend IsGlobalFunc As Boolean
        Friend DontExecuteSub As Boolean
        Friend KeepReturnValue As Boolean

        Public Overrides Function GetStatementAt(lineNumber As Integer) As Statement
            If lineNumber < StartToken.Line Then Return Nothing
            If Args.Count > 0 AndAlso lineNumber <= Args.Last.EndToken.Line Then Return Me
            Return Nothing
        End Function

        Public Overrides Sub AddSymbols(symbolTable As SymbolTable)
            MyBase.AddSymbols(symbolTable)
            Name.Parent = Me

            For Each arg In Args
                arg.Parent = Me
                arg.AddSymbols(symbolTable)
            Next

            Name.SymbolType = CompletionItemType.SubroutineName
            symbolTable.AddIdentifier(Name)
        End Sub

        Public Overrides Sub EmitIL(scope As CodeGenScope)
            If scope.ForGlobalHelp Then Return

            If Args.Count > 0 Then
                For Each arg In Args
                    arg.EmitIL(scope)
                Next
            End If

            Dim methodInfo = scope.MethodBuilders(Name.LCaseText)
            scope.ILGenerator.EmitCall(OpCodes.Call, methodInfo, Nothing)
            If Not KeepReturnValue AndAlso methodInfo.ReturnType IsNot GetType(Void) Then
                scope.ILGenerator.Emit(OpCodes.Pop)
            End If
        End Sub

        Public Overrides Sub PopulateCompletionItems(
                        bag As CompletionBag,
                        line As Integer,
                        column As Integer,
                        globalScope As Boolean
                   )

            Dim subName = Name.Text
            Dim subName2 = subName.ToLower
            If Name.Contains(line, column, True) Then
                bag.CompletionItems.Add(
                    New CompletionItem() With {
                       .DisplayName = subName,
                       .ItemType = CompletionItemType.SubroutineName,
                       .Key = subName,
                       .ReplacementText = subName,
                       .DefinitionIdintifier = (From subroutine In bag.SymbolTable.Subroutines
                                                Where subroutine.Key = subName2).FirstOrDefault.Value
                    }
                )

            ElseIf Args IsNot Nothing AndAlso Args.Count > 0 AndAlso line >= Args(0).StartToken.Line Then
                CompletionHelper.FillExpressionItems(bag)
                CompletionHelper.FillSubroutines(bag, True)
            End If
        End Sub


        Public Overrides Function ToString() As String
            Dim sb As New Text.StringBuilder(Name.Text)
            Dim n = If(Args Is Nothing, -1, Args.Count - 1)

            sb.Append("(")
            For i = 0 To n
                sb.Append(Args(0).ToString())
                If i < n Then sb.Append(", ")
            Next
            sb.AppendLine(")")

            Return sb.ToString()
        End Function

        Public Overrides Function Execute(runner As ProgramRunner) As Statement
            Dim subName = Name.LCaseText
            Dim targetRunner = If(IsGlobalFunc, runner.GlobalRunner, runner)
            If OuterSubroutine Is Nothing Then
                OuterSubroutine = SubroutineStatement.GetSubroutine(Me)
            End If

            Dim parentSubName = OuterSubroutine?.Name.LCaseText
            Dim recursive = (subName = parentSubName AndAlso targetRunner Is runner)
            Dim locals As Dictionary(Of String, Library.Primitive)

            If recursive Then
                Dim keyPrefix = subName & "."
                locals = New Dictionary(Of String, Library.Primitive)

                For Each item In targetRunner.Fields
                    If item.Key.StartsWith(keyPrefix) Then
                        locals(item.Key) = item.Value
                    End If
                Next
            End If

            Dim subToken = targetRunner.SymbolTable.Subroutines(subName)
            If DontExecuteSub And subToken.Type = TokenType.Sub Then
                Return Nothing
            Else
                DontExecuteSub = False
            End If

            If IsGlobalFunc Then
                If runner.Evaluating Then
                    targetRunner = targetRunner.GetEvaluationRunner()
                Else
                    targetRunner.HasBeenPaused = False
                    targetRunner.StepAround = False
                    targetRunner.Depth = 0
                    targetRunner.DebuggerState = DebuggerState.Running
                    If runner.DebuggerCommand = DebuggerCommand.StepInto OrElse runner.DebuggerCommand = DebuggerCommand.StopOnNextLine Then
                        targetRunner.DebuggerCommand = DebuggerCommand.StepInto
                        runner.DebuggerCommand = DebuggerCommand.Run
                    Else
                        targetRunner.DebuggerCommand = DebuggerCommand.Run
                    End If
                    runner.Engine.CurrentRunner = targetRunner
                End If
            End If

            targetRunner.runnerThread = runner.runnerThread
            Dim subStatement As SubroutineStatement = subToken.Parent
            subStatement.SetParams(targetRunner, Args)

            Dim result = subStatement.Execute(targetRunner)

            subStatement.ResetLocals(targetRunner)
            If recursive Then ' restore locals
                For Each item In locals
                    targetRunner.Fields(item.Key) = item.Value
                Next
            End If

            If IsGlobalFunc Then
                If Not runner.Evaluating Then
                    Dim dc = targetRunner.DebuggerCommand
                    If dc <> DebuggerCommand.Run Then
                        runner.DebuggerCommand = DebuggerCommand.StepInto
                        runner.StepAround = False
                        runner.Depth = 0
                    ElseIf targetRunner.HasBeenPaused Then
                        runner.DebuggerCommand = dc
                    End If

                    targetRunner.ChangeDebuggerState(DebuggerState.Finished)
                    targetRunner.ContinueAll()
                End If
            End If

            runner.PauseAtReturn = targetRunner.HasBeenPaused
            runner.Engine.CurrentRunner = runner
            Return If(TypeOf result Is EndDebugging, result, Nothing)
        End Function

        Public Overrides Function ToVB(symbolTable As SymbolTable) As String
            Dim sb As New Text.StringBuilder(Name.Text)
            Dim n = If(Args Is Nothing, -1, Args.Count - 1)

            sb.Append("(")
            For i = 0 To n
                sb.Append(Args(0).ToVB())
                If i < n Then sb.Append(", ")
            Next
            sb.AppendLine(")")

            Return sb.ToString()
        End Function
    End Class
End Namespace
