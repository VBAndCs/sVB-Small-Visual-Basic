Imports System.Threading
Imports Microsoft.SmallVisualBasic.Statements

Namespace Microsoft.SmallVisualBasic.Engine

    Public Class ProgramRunner
        Public Property Breakpoints As New List(Of Integer)
        Public Property DebuggerCommand As DebuggerCommand = DebuggerCommand.Run
        Public Property DebuggerState As DebuggerState
        Public Property SubroutineInstructions As Dictionary(Of String, List(Of Instruction))
        Public Property Engine As ProgramEngine
        Public Property Fields As New Dictionary(Of String, Library.Primitive)

        Friend PreviousLineNumber As Integer = -1
        Public CurrentLineNumber As Integer
        Friend BreakMode As Boolean
        Public DocLineOffset As Integer
        Friend SymbolTable As SymbolTable
        Public TypeInfoBag As TypeInfoBag
        Public Parser As Parser
        Friend EventThreads As New Collections.Concurrent.ConcurrentDictionary(Of String, Boolean)()

        Friend Function CreateSubRunner() As ProgramRunner
            Dim runner As New ProgramRunner(_Engine, Parser, True) With {
                ._Breakpoints = _Breakpoints,
                ._DebuggerCommand = _DebuggerCommand,
                ._DebuggerState = _DebuggerState,
                .Fields = Fields,
                .isSubRunner = True
            }
            _Engine.SubRunners(Parser) = runner
            Return runner
        End Function

        Public Sub New(engine As ProgramEngine, parser As Parser, Optional isSubRunner As Boolean = False)
            DocLineOffset = parser.DocStartLine
            _Engine = engine
            TypeInfoBag = Compiler.TypeInfoBag
            Me.Parser = parser
            SymbolTable = Me.Parser.SymbolTable
            If Not isSubRunner Then _Engine.Runners(Me.Parser) = Me
        End Sub

        Dim isGlobalInitialized As Boolean = False
        Friend HasBeenPaused As Boolean

        Friend ReadOnly Property GlobalRunner As ProgramRunner
            Get
                Dim gRunner = _Engine.Runners(_Engine.GlobalParser)
                If Not gRunner.isGlobalInitialized Then
                    gRunner.isGlobalInitialized = True
                    gRunner.StepAround = False
                    gRunner.Depth = 0

                    If DebuggerCommand = DebuggerCommand.StepInto OrElse DebuggerCommand = DebuggerCommand.StopOnNextLine Then
                        gRunner.DebuggerCommand = DebuggerCommand.StepInto
                        DebuggerCommand = DebuggerCommand.Run
                    Else
                        gRunner.DebuggerCommand = DebuggerCommand.Run
                    End If

                    gRunner.CurrentThread = CurrentThread
                    _Engine.CurrentRunner = gRunner
                    gRunner.HasBeenPaused = False

                    gRunner.Execute(_Engine.GlobalParser.ParseTree)

                    Dim dc = gRunner.DebuggerCommand
                    If dc <> DebuggerCommand.Run Then
                        DebuggerCommand = DebuggerCommand.StepInto
                    ElseIf gRunner.HasBeenPaused Then
                        DebuggerCommand = dc
                    End If

                    PauseAtReturn = gRunner.HasBeenPaused
                    _Engine.CurrentRunner = Me
                End If
                Return gRunner
            End Get
        End Property

        Public Function EvaluateExpression(ByRef text As String, subName As Token) As Library.Primitive?
            Dim reader As New IO.StringReader(text)
            Dim lines = New List(Of String)

            Do
                Dim line = reader.ReadLine
                If line Is Nothing Then Exit Do
                lines.Add(line)
            Loop

            Dim tokens = LineScanner.GetTokens(lines(0), 0, lines)
            Dim subSt As SubroutineStatement = Nothing

            If Not subName.IsIllegal Then
                subSt = New SubroutineStatement() With {.Name = subName}
                Dim subToken As New Token() With {
                    .Line = subName.Line,
                    .Column = 0,
                    .Parent = subSt
                }
                subSt.StartToken = subToken
                subSt.SubToken = subToken

                For i = 0 To tokens.Count - 1
                    Dim token = tokens(i)
                    token.Parent = subSt
                    tokens(i) = token
                Next
            End If

            SubroutineStatement.Current = subSt
            Dim expr = Parser.BuildExpression(New TokenEnumerator(tokens), True)
            SubroutineStatement.Current = Nothing
            If expr Is Nothing Then Return Nothing

            ' Set the display text (by ref)
            text = expr.ToString()

            Dim mc = TryCast(expr, Expressions.MethodCallExpression)
            If mc IsNot Nothing Then
                If mc.TypeName.Text = "" Then
                    Dim subroutines = Parser.SymbolTable.Subroutines
                    Dim key = mc.MethodName.LCaseText
                    If subroutines.ContainsKey(key) Then
                        Dim subroutine = CType(subroutines(key).Parent, SubroutineStatement)
                        If subroutine.SubToken.Type = TokenType.Sub Then
                            Return "A subroutine call doesn't return any value!"
                        End If
                    End If
                End If
            End If

            Return expr.Evaluate(Me)
        End Function

        Friend ReadOnly Property ShortStep As Boolean
            Get
                Return DebuggerCommand = DebuggerCommand.ShortStepOver OrElse
                            DebuggerCommand = DebuggerCommand.ShortStepOut
            End Get
        End Property

        Friend Sub RunForm(formName As String)
            Dim className = (WinForms.Forms.FormPrefix & formName).ToLower()
            Dim formParser As Parser
            For Each p In _Engine.Parsers
                If p.ClassName.ToLower() = className Then
                    formParser = p
                    Exit For
                End If
            Next

            If formParser Is Nothing Then
                Throw New Exception("There is no form name {formName} in the project")
            End If

            Dim formRunner As ProgramRunner
            If _Engine.Runners.ContainsKey(formParser) Then
                formRunner = _Engine.Runners(formParser)
            Else
                formRunner = New ProgramRunner(_Engine, formParser)
            End If

            formRunner.HasBeenPaused = False
            formRunner.StepAround = False
            formRunner.Depth = 0
            If DebuggerCommand = DebuggerCommand.StepInto OrElse DebuggerCommand = DebuggerCommand.StopOnNextLine Then
                formRunner.DebuggerCommand = DebuggerCommand.StepInto
                DebuggerCommand = DebuggerCommand.Run
            Else
                formRunner.DebuggerCommand = DebuggerCommand.Run
            End If

            formRunner.CurrentThread = CurrentThread
            _Engine.CurrentRunner = formRunner
            formRunner.Execute(formParser.ParseTree)

            Dim dc = formRunner.DebuggerCommand
            If dc <> DebuggerCommand.Run Then
                DebuggerCommand = DebuggerCommand.StepInto
            ElseIf formRunner.HasBeenPaused Then
                DebuggerCommand = dc
            End If

            PauseAtReturn = formRunner.hasBeenPaused
            _Engine.CurrentRunner = Me
        End Sub

        Friend Sub SetGlobalField(name As String, value As Library.Primitive)
            GlobalRunner.Fields(name) = value
        End Sub

        Friend Function GetGlobalField(name As String) As Library.Primitive?
            If GlobalRunner.Fields.ContainsKey(name) Then
                Return GlobalRunner.Fields(name)
            Else
                Return Nothing
            End If
        End Function

        Public Function GetKey(identifier As Token) As String
            Dim variableName = identifier.LCaseText
            Dim Subroutine = identifier.SubroutineName?.ToLower()
            If Subroutine = "" Then Return variableName

            Dim key = $"{Subroutine}.{variableName}"
            If SymbolTable.LocalVariables.ContainsKey(key) Then
                Return key
            Else
                Return variableName
            End If
        End Function

        Friend Sub ChangeDebuggerState(state As DebuggerState)
            If state = DebuggerState.Running Then
                isPaused = False
                If _DebuggerState = state Then Return
                Library.Program.ActivateWindow()

            ElseIf state = DebuggerState.Finished Then
                isPaused = False
                If _DebuggerState = state OrElse _DebuggerState = DebuggerState.Running Then
                    _DebuggerState = state
                    Return
                End If
                Library.Program.ActivateWindow()
            End If

            _DebuggerState = state
            _Engine.RaiseDebuggerStateChangedEvent(Me)
        End Sub

        Friend Sub CheckForExecutionBreakAtLine(lineNumber As Integer, Optional isSub As Boolean = False)
            If StepAround Then
                If Depth > 0 OrElse (isSub AndAlso lineNumber <> StepLineNumber) Then
                    Return
                End If
            End If

            CurrentLineNumber = lineNumber
            CheckForExecutionBreak(False, Nothing)
            PreviousLineNumber = lineNumber
        End Sub

        Private Sub CheckForExecutionBreak(isFirstStatement As Boolean, CurrentStatement As Statement)
            If _Engine.Evaluating Then Return

            If isFirstStatement AndAlso _Engine.StopOnFirstStaement Then
                _Engine.StopOnFirstStaement = False
                Pause()

            ElseIf TypeOf CurrentStatement Is StopStatement Then
                DebuggerCommand = DebuggerCommand.StepInto
                Pause()

            ElseIf CurrentLineNumber <> PreviousLineNumber Then
                If Breakpoints.Contains(CurrentLineNumber - DocLineOffset) OrElse
                        DebuggerCommand = DebuggerCommand.StepInto OrElse
                        DebuggerCommand = DebuggerCommand.StepOver OrElse
                        DebuggerCommand = DebuggerCommand.ShortStepOver Then
                    Pause()
                ElseIf DebuggerCommand = DebuggerCommand.StopOnNextLine Then
                    DebuggerCommand = DebuggerCommand.Run
                    Pause()
                End If
            End If
        End Sub


        Friend CurrentThread As Thread

        Friend Sub [Continue](Optional all As Boolean = False)
            forcedToBreak = False
            ' This maybe a subrunner, so we must resum it
            Try
                CurrentThread?.Resume()
            Catch
            End Try

            If all Then ContinueAll()
        End Sub

        Dim isPaused As Boolean
        Dim forcedToBreak As Boolean

        Private Sub Pause(Optional isError As Boolean = False)
            If Library.Program.IsTerminated Then Return
            BreakMode = True
            StepAround = False
            isPaused = True

            Dim mainRunner = _Engine.MainRunner
            If Me IsNot mainRunner AndAlso Not mainRunner.forcedToBreak AndAlso
                    mainRunner.DebuggerState = DebuggerState.Running Then
                mainRunner.forcedToBreak = True
                If mainRunner.DebuggerState = DebuggerState.Running Then
                    mainRunner.DebuggerCommand = DebuggerCommand.StepInto
                    mainRunner.StepAround = False
                    mainRunner.Depth = 0
                End If
            End If

            If Me IsNot mainRunner OrElse Not forcedToBreak Then
                HasBeenPaused = True
                ChangeDebuggerState(If(isError, DebuggerState.Error, DebuggerState.Paused))
                _Engine.PausedRunner = Me
            End If

            Try
                DebuggerCommand = DebuggerCommand.Run ' After resuming
                CurrentThread.Suspend()
            Catch ex As Exception
            End Try
        End Sub

        Public Sub Reset()
            Fields.Clear()
            Depth = 0
            StepAround = False
        End Sub

        Friend Depth As Integer = 0
        Friend StepAround As Boolean = False
        Friend StepLineNumber As Integer = -1
        Friend PauseAtReturn As Boolean

        Friend Sub ContinueAll()
            For Each runner In _Engine.Runners.Values
                If runner IsNot Me AndAlso runner.ForcedToBreak AndAlso
                            runner.CurrentThread IsNot Nothing AndAlso
                            runner.CurrentThread.IsAlive Then
                    runner.forcedToBreak = False
                    Try
                        runner.CurrentThread.Resume()
                    Catch
                    End Try
                End If
                runner.BreakMode = False
            Next
        End Sub

        Friend Sub StepInto()
            DebuggerCommand = DebuggerCommand.StepInto
            Depth = 0
            StepAround = False
            [Continue]()
        End Sub

        Friend Sub StepOut()
            DebuggerCommand = DebuggerCommand.StepOut
            UpdateMarkers()
            StepLineNumber = CurrentLineNumber
            Depth = 1
            StepAround = True
            [Continue]()
        End Sub

        Friend Sub ShortStepOut()
            DebuggerCommand = DebuggerCommand.ShortStepOut
            UpdateMarkers()
            StepLineNumber = CurrentLineNumber
            Depth = 1
            StepAround = True
            [Continue]()
        End Sub

        Private Sub UpdateMarkers()
            ' Remove the break marker in case of long running code
            _DebuggerState = DebuggerState.Running
            _Engine.RaiseDebuggerStateChangedEvent(Me)
        End Sub

        Friend Sub StepOver()
            If StepAround Then Return
            DebuggerCommand = DebuggerCommand.StepOver
            UpdateMarkers()
            Depth = 0
            StepAround = True
            StepLineNumber = CurrentLineNumber
            [Continue]()
        End Sub

        Friend Sub ShortStepOver()
            DebuggerCommand = DebuggerCommand.ShortStepOver
            UpdateMarkers()
            Depth = 0
            StepAround = True
            StepLineNumber = CurrentLineNumber
            [Continue]()
        End Sub

        Public Exception As Exception
        Friend ResumeTrials As Integer
        Private isSubRunner As Boolean
        Friend Shared EventsCommand As DebuggerCommand

        Friend Function Execute(statements As List(Of Statement)) As Statement
            PreviousLineNumber = -1
            If statements.Count = 0 Then Return Nothing

            'CurrentThread = Threading.Thread.CurrentThread
            Dim startLine = statements(0).StartToken.Line
            Dim endLine = statements.Last.StartToken.Line
            Dim isFirstStatement = True

            For i = 0 To statements.Count - 1
                Dim st = statements(i)
                If TypeOf st Is SubroutineStatement OrElse
                    TypeOf st Is EmptyStatement OrElse
                    TypeOf st Is IllegalStatement Then Continue For

                _Engine.CurrentRunner = Me
                CurrentLineNumber = st.StartToken.Line
                Dim isNotGenCode = CurrentLineNumber - DocLineOffset > -1

                If isNotGenCode Then
                    If Not StepAround OrElse Depth = 0 Then
                        CheckForExecutionBreak(isFirstStatement, st)
                    End If
                End If

                PreviousLineNumber = CurrentLineNumber
                Dim result As Statement
                Try
                    result = st.Execute(Me)
                    Exception = Library.Program.Exception
                    If Exception IsNot Nothing Then
                        CurrentThread = Thread.CurrentThread
                        Pause(isError:=True)
                        Library.Program.End()
                        Return New EndDebugging()
                    End If

                Catch ex As Exception
                    Exception = ex
                    CurrentThread = Thread.CurrentThread
                    Pause(isError:=True)
                    Library.Program.End()
                    Return New EndDebugging()
                End Try

                If Depth = 0 Then
                    If DebuggerCommand = DebuggerCommand.ShortStepOver Then
                        StepAround = False
                    ElseIf DebuggerCommand = DebuggerCommand.ShortStepOut Then
                        DebuggerCommand = DebuggerCommand.StopOnNextLine
                        StepAround = False
                    End If
                End If

                If isNotGenCode Then isFirstStatement = False

                If TypeOf result Is EndDebugging Then
                    Return result
                ElseIf SmallVisualBasic.Library.Program.IsTerminated Then
                    Return New EndDebugging()
                End If

                If PauseAtReturn Then
                    PauseAtReturn = False
                    If isNotGenCode Then
                        PreviousLineNumber = -1
                        CurrentLineNumber = st.StartToken.Line
                        If Not StepAround OrElse Depth = 0 Then
                            Select Case DebuggerCommand
                                Case DebuggerCommand.Run, DebuggerCommand.StepOver, DebuggerCommand.ShortStepOver
                                    ' don't stop at caller
                                Case Else
                                    CheckForExecutionBreak(isFirstStatement, st)
                            End Select
                        End If
                        PreviousLineNumber = CurrentLineNumber
                    End If
                End If

                If TypeOf result Is ReturnStatement OrElse TypeOf result Is JumpLoopStatement Then
                    Return result
                ElseIf TypeOf result Is GotoStatement Then
                    Dim gotoLine = CType(result, GotoStatement).Label.Line
                    If gotoLine < startLine OrElse gotoLine > endLine Then
                        Return result
                    Else
                        For j = 0 To statements.Count - 1
                            If statements(j).StartToken.Line = gotoLine Then
                                i = j - 1
                                Exit For
                            End If
                        Next
                    End If
                End If
            Next
            Return Nothing
        End Function

        Friend Sub DecreaseDepthOfShortStepOut(ByRef mustStepOut As Boolean)
            If Depth < 1 Then Return

            If DebuggerCommand = DebuggerCommand.ShortStepOut Then
                If Depth = 1 Then
                    mustStepOut = True
                    DebuggerCommand = DebuggerCommand.ShortStepOver
                ElseIf Not mustStepOut Then
                    Depth -= 1
                End If
            End If
        End Sub

        Friend Sub IncreaseDepthOfShortSteps(ByRef mustStepOut As Boolean)
            Dim isShortStepOut = (DebuggerCommand = DebuggerCommand.ShortStepOut)

            If isShortStepOut AndAlso StepLineNumber = CurrentLineNumber Then
                Depth = 2
                mustStepOut = True
                'DebuggerCommand = DebuggerCommand.ShortStepOver

            ElseIf Not mustStepOut AndAlso (isShortStepOut OrElse
                    DebuggerCommand = DebuggerCommand.ShortStepOver) Then
                mustStepOut = True
                Depth += 1
            End If
        End Sub

    End Class

End Namespace
