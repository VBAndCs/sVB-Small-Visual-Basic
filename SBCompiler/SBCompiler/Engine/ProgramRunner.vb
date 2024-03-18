Imports System.Threading
Imports Microsoft.SmallVisualBasic.Statements

Namespace Microsoft.SmallVisualBasic.Engine

    Public Class ProgramRunner
        Friend PreviousLineNumber As Integer = -1
        Friend CurrentLineNumber As Integer

        Public Property Breakpoints As New List(Of Integer)

        Public Property CurrentStatement As Statement

        Public Property DebuggerCommand As DebuggerCommand

        Public Property DebuggerState As DebuggerState

        Public Property Instructions As List(Of Instruction)

        Public Property LabelMap As New Dictionary(Of String, Integer)

        Public Property SubroutineInstructions As Dictionary(Of String, List(Of Instruction))

        Public DocLineOffset As Integer

        Public Property Engine As ProgramEngine
        Public Property Fields As New Dictionary(Of String, Library.Primitive)

        Friend SymbolTable As SymbolTable
        Public TypeInfoBag As TypeInfoBag
        Public currentParser As Parser

        Public Sub New(engine As ProgramEngine, parser As Parser)
            DocLineOffset = parser.DocStartLine
            _Engine = engine
            TypeInfoBag = Compiler.TypeInfoBag
            currentParser = parser
            SymbolTable = currentParser.SymbolTable
            _Engine.runners(currentParser) = Me
            _Engine.CurrentDebuggerState = _DebuggerState
            _Engine.RaiseDebuggerStateChangedEvent()
        End Sub

        Dim isGlobalInitialized As Boolean = False
        Dim hasBeenPaused As Boolean

        Friend ReadOnly Property GlobalRunner As ProgramRunner
            Get
                Dim gRunner = _Engine.Runners(_Engine.GlobalParser)
                If Not gRunner.isGlobalInitialized Then
                    gRunner.isGlobalInitialized = True
                    If DebuggerCommand = DebuggerCommand.StepInto Then
                        gRunner.DebuggerCommand = DebuggerCommand.StepInto
                    End If
                    gRunner.CurrentThread = Threading.Thread.CurrentThread
                    gRunner.hasBeenPaused = False
                    gRunner.Execute(_Engine.GlobalParser.ParseTree)
                    PauseAtReturn = gRunner.hasBeenPaused
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
            Dim expr = currentParser.BuildExpression(New TokenEnumerator(tokens), True)
            If expr Is Nothing Then Return Nothing
            text = expr.ToString()
            Return expr.Evaluate(Me)
        End Function

        Public ReadOnly Property ShortStep As Boolean
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

            formRunner.hasBeenPaused = False
            formRunner.Execute(formParser.ParseTree)
            PauseAtReturn = formRunner.hasBeenPaused
            _Engine.CurrentRunner = Me
        End Sub

        Friend Sub SetGlobalField(name As String, value As Library.Primitive)
            GlobalRunner.Fields(name) = value
        End Sub

        Friend Function GetGlobalField(name As String) As Library.Primitive
            If GlobalRunner.Fields.ContainsKey(name) Then
                Return GlobalRunner.Fields(name)
            Else
                Return 0
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

        Private Sub ChangeDebuggerState(state As DebuggerState)
            If _DebuggerState <> state Then
                _DebuggerState = state
                _Engine.CurrentDebuggerState = state
                _Engine.RaiseDebuggerStateChangedEvent()
            End If
        End Sub

        Friend Sub CheckForExecutionBreakAtLine(lineNumber As Integer, Optional isSub As Boolean = False)
            If StepAround Then
                If Depth > 0 OrElse (isSub AndAlso lineNumber <> StepOverLineNumber) Then
                    Return
                End If
            End If

            CurrentLineNumber = lineNumber
            CheckForExecutionBreak(False)
            PreviousLineNumber = lineNumber
        End Sub

        Private Sub CheckForExecutionBreak(isFirstStatement As Boolean)
            If isFirstStatement AndAlso _Engine.StopOnFirstStaement Then
                _Engine.StopOnFirstStaement = False
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

        Public Sub [Continue]()
            isPaused = False
            ChangeDebuggerState(DebuggerState.Running)

            If CurrentThread?.IsAlive Then
                Try
                    CurrentThread.Resume()
                Catch
                End Try
            End If
        End Sub

        Dim isPaused As Boolean

        Public Sub Pause()
            hasBeenPaused = True
            StepAround = False
            _Engine.BreakMode = True
            isPaused = True
            ChangeDebuggerState(DebuggerState.Paused)
            Try
                CurrentThread.Suspend()
            Catch
            End Try
        End Sub

        Public Sub Reset()
            Fields.Clear()
            Depth = 0
        End Sub

        Friend Depth As Integer = 0
        Friend StepAround As Boolean = False
        Friend StepOverLineNumber As Integer = -1
        Friend EventTreads As New Collections.Concurrent.ConcurrentDictionary(Of Reflection.EventInfo, Boolean)
        Friend PauseAtReturn As Boolean

        Public Sub StepInto()
            DebuggerCommand = DebuggerCommand.StepInto
            Depth = 0
            StepAround = False
            [Continue]()
        End Sub

        Friend Sub StepOut()
            DebuggerCommand = DebuggerCommand.StepOut
            StepOverLineNumber = CurrentLineNumber
            Depth = 1
            StepAround = True
            [Continue]()
        End Sub

        Friend Sub ShortStepOut()
            DebuggerCommand = DebuggerCommand.ShortStepOut
            StepOverLineNumber = CurrentLineNumber
            Depth = 1
            StepAround = True
            [Continue]()
        End Sub

        Public Sub StepOver()
            If StepAround Then Return

            DebuggerCommand = DebuggerCommand.StepOver
            Depth = 0
            StepAround = True
            StepOverLineNumber = CurrentLineNumber
            [Continue]()
        End Sub

        Friend Sub ShortStepOver()
            DebuggerCommand = DebuggerCommand.ShortStepOver
            Depth = 0
            StepAround = True
            StepOverLineNumber = CurrentLineNumber
            [Continue]()
        End Sub

        Friend Function Execute(statements As List(Of Statement)) As Statement
            PreviousLineNumber = -1
            If statements.Count = 0 Then Return Nothing

            Dim startLine = statements(0).StartToken.Line
            Dim endLine = statements.Last.StartToken.Line
            _Engine.CurrentRunner = Me
            Dim isFirstStatement = True

            For i = 0 To statements.Count - 1
                Dim st = statements(i)
                If TypeOf st Is SubroutineStatement OrElse
                    TypeOf st Is EmptyStatement OrElse
                    TypeOf st Is IllegalStatement Then Continue For

                CurrentStatement = st
                CurrentLineNumber = CurrentStatement.StartToken.Line
                Dim isNotGenCode = CurrentLineNumber - DocLineOffset > -1

                If isNotGenCode Then
                    If Not StepAround OrElse Depth = 0 Then
                        CheckForExecutionBreak(isFirstStatement)
                    End If
                End If

                PreviousLineNumber = CurrentLineNumber
                Dim result = st.Execute(Me)

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
                        CurrentStatement = st
                        PreviousLineNumber = -1
                        CurrentLineNumber = CurrentStatement.StartToken.Line
                        If Not StepAround OrElse Depth = 0 Then
                            Select Case DebuggerCommand
                                Case DebuggerCommand.Run, DebuggerCommand.StepOver, DebuggerCommand.ShortStepOver
                                    ' don't stop at caller
                                Case Else
                                    CheckForExecutionBreak(isFirstStatement)
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

            If isShortStepOut AndAlso StepOverLineNumber = CurrentLineNumber Then
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
