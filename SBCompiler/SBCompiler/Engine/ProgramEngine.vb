Imports System
Imports System.Collections.Generic
Imports System.Reflection
Imports System.Threading
Imports Microsoft.SmallVisualBasic.Expressions
Imports Microsoft.SmallVisualBasic.Library

Namespace Microsoft.SmallVisualBasic.Engine
    Public Class ProgramEngine

        Public Event DebuggerStateChanged(runner As ProgramRunner)

        Public ReadOnly Property Parsers As List(Of Parser)

        Public ReadOnly Property CurrentLineNumber As Integer
            Get
                Return CurrentRunner.CurrentLineNumber
            End Get
        End Property

        Dim _CurrentRunner As ProgramRunner

        Public Property CurrentRunner As ProgramRunner
            Get
                Return _CurrentRunner
            End Get

            Set(runner As ProgramRunner)
                If runner Is Nothing OrElse Not runner.Evaluating Then
                    _CurrentRunner = runner
                End If
            End Set
        End Property

        Public Property PausedRunner As ProgramRunner

        Public Runners As Dictionary(Of Parser, ProgramRunner)
        Public SubRunners As Dictionary(Of Parser, ProgramRunner)

        Public Sub New(parsers As List(Of Parser))
            Me.Parsers = parsers
            InitializeRunner()
        End Sub

        Friend GlobalParser As Parser

        Public ReadOnly Property GlobalRunner() As ProgramRunner
            Get
                Return CurrentRunner.GlobalRunner
            End Get
        End Property

        Private ReadOnly Property CanRun As Boolean
            Get
                If _PausedRunner Is Nothing Then
                    For Each runner In Runners.Values
                        If runner.isPaused Then
                            _PausedRunner = runner
                            Return True
                        End If
                    Next
                    Return False
                Else
                    Return True
                End If
            End Get
        End Property

        Private Sub InitializeRunner()
            Dim mainParser, formParsr As Parser

            For Each p In Parsers
                If p.IsMainForm Then
                    mainParser = p
                    Exit For
                ElseIf p.IsGlobal Then
                    GlobalParser = p
                Else
                    formParsr = p
                End If
            Next

            If mainParser Is Nothing Then mainParser = formParsr
            Runners = New Dictionary(Of Parser, ProgramRunner)
            SubRunners = New Dictionary(Of Parser, ProgramRunner)
            MainRunner = New ProgramRunner(Me, mainParser)
            _CurrentRunner = MainRunner
            If GlobalParser IsNot Nothing Then
                Dim __ = New ProgramRunner(Me, GlobalParser)
            End If
        End Sub

        Friend Sub RaiseDebuggerStateChangedEvent(runner As ProgramRunner)
            RaiseEvent DebuggerStateChanged(If(runner, CurrentRunner))
        End Sub

        Public Sub AddBreakpoint(lineNumber As Integer)
            CurrentRunner.Breakpoints.Add(lineNumber)
        End Sub

        Public Sub [Continue]()
            If CanRun Then
                _PausedRunner.DebuggerCommand = DebuggerCommand.Run
                _PausedRunner.ChangeDebuggerState(DebuggerState.Running)
                _PausedRunner.Continue(True)
            End If
        End Sub

        Public Sub Pause()
            ' Don't call pause directly, because it can be executed after the thread is ended.
            ' Call StepInto, because it changes the command while the thread is running, which forces it to pause at the next statement
            If SubRunners.ContainsKey(_CurrentRunner.Parser) Then
                ' Timers handlers change their states rapidily
                Dim st = Now
                Do
                    If _CurrentRunner.DebuggerState = DebuggerState.Running Then
                        CurrentRunner.StepInto()
                        Return
                    End If
                    Thread.Sleep(10)
                    If (Now - st).TotalMilliseconds > 1000 Then Return
                Loop

            ElseIf _CurrentRunner.DebuggerState = DebuggerState.Running Then
                CurrentRunner.StepInto()
            End If
        End Sub

        Public Sub StepInto()
            If CanRun Then
                _PausedRunner.StepInto()
            End If
        End Sub

        Public Sub StepOut()
            If CanRun Then
                _PausedRunner.StepOut()
            End If
        End Sub

        Public Sub ShortStepOut()
            If CanRun Then
                _PausedRunner.ShortStepOut()
            End If
        End Sub

        Public Sub StepOver()
            If CanRun Then
                _PausedRunner.StepOver()
            End If
        End Sub

        Public Sub ShortStepOver()
            If CanRun Then
                _PausedRunner.ShortStepOver()
            End If
        End Sub

        Public Sub Reset()
            CurrentRunner.Reset()
        End Sub


        Friend StopOnFirstStaement As Boolean

        Public Property BreakMode As Boolean

        Public Sub RunProgram(Optional stopOnFirstStaement As Boolean = False)
            Me.StopOnFirstStaement = stopOnFirstStaement
            DoRunProgram()
        End Sub

        Friend MainRunner As ProgramRunner

        Private Sub DoRunProgram()
            Dim debuggerThread As New Thread(
                Sub()
                    WinForms.GeometricPath.CreatePath()
                    Try
                        Library.TextWindow.ClearIfLoaded()
                    Catch
                    End Try

                    Library.Program.IsTerminated = False
                    Library.Program.Exception = Nothing
                    MainRunner.CurrentThread = Thread.CurrentThread
                    Dim result = MainRunner.Execute(MainRunner.Parser.ParseTree)
                    If Library.Program.IsTerminated Then Return

                    MainRunner.ChangeDebuggerState(DebuggerState.Running)
                    If TypeOf result Is Statements.EndDebugging Then Return

                    ' Allow event handlers to break
                    If MainRunner.DebuggerCommand = DebuggerCommand.Run Then
                        ProgramRunner.EventsCommand = DebuggerCommand.Run
                    Else
                        ProgramRunner.EventsCommand = DebuggerCommand.StepInto
                    End If

                    MainRunner.Depth = 0
                    MainRunner.StepAround = False
                    MainRunner.DebuggerState = DebuggerState.Finished

                    Try
                        Library.TextWindow.PauseThenClose()
                    Catch
                    End Try
                End Sub
             )
            debuggerThread.IsBackground = True
            debuggerThread.Start()
        End Sub

        Private Sub [End]()
            For Each runner In Runners.Values
                Abort(runner.CurrentThread)
            Next

            For Each runner In SubRunners.Values
                Abort(runner.CurrentThread)
            Next

            For Each parser In Parsers
                Try
                    Abort(parser.EvaluationRunner?.CurrentThread)
                Catch ex As Exception
                End Try
            Next
        End Sub

        Private Sub Abort(thread As Thread)
            If thread Is Nothing Then Return

            If (thread.ThreadState And ThreadState.Aborted) <> 0 OrElse
                   (thread.ThreadState And ThreadState.AbortRequested) <> 0 OrElse
                   (thread.ThreadState And ThreadState.Unstarted) <> 0 Then
                Return
            ElseIf (thread.ThreadState And ThreadState.Suspended) <> 0 OrElse
                    (thread.ThreadState And ThreadState.SuspendRequested) <> 0 Then
                Try
                    thread.Resume()
                Catch ex As Exception
                End Try
            End If

            Try
                thread.Abort()
            Catch ex As Exception
            End Try
        End Sub

        Friend Sub ResetEvaluationRunners()
            For Each parser In Parsers
                parser.EvaluationRunner = Nothing
            Next
        End Sub

        Public Sub Dispose()
            [End]()
            Me.Parsers.Clear()
            Me.CurrentRunner = Nothing
            Me.GlobalParser = Nothing
            Me.Runners.Clear()
            Me.SubRunners.Clear()
        End Sub

        Public Function GetEvaluationRunner(file As String) As ProgramRunner
            For Each r In Runners.Values
                If r.Parser.DocPath.ToLower = file.ToLower() Then
                    Return r.GetEvaluationRunner()
                End If
            Next
            Return Nothing
        End Function


        Public Sub UpdateBreakpoints(breakpoints As List(Of Integer), file As String)
            file = file.ToLower()

            For Each runner In Runners.Values
                If runner.Parser.DocPath.ToLower() = file Then
                    runner.Breakpoints = breakpoints
                End If
            Next

            For Each runner In SubRunners.Values
                If runner.Parser.DocPath.ToLower() = file Then
                    runner.Breakpoints = breakpoints
                End If
            Next
        End Sub
    End Class

End Namespace

