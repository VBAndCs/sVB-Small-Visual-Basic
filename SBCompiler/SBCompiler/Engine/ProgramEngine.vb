Imports System
Imports System.Collections.Generic
Imports System.Reflection
Imports System.Threading
Imports Microsoft.SmallVisualBasic.Expressions
Imports Microsoft.SmallVisualBasic.Library

Namespace Microsoft.SmallVisualBasic.Engine
    Public Class ProgramEngine

        Public Event DebuggerStateChanged(runner As ProgramRunner)

        Public Property Parsers As List(Of Parser)

        Public ReadOnly Property CurrentLineNumber As Integer
            Get
                Return CurrentRunner.CurrentLineNumber
            End Get
        End Property


        Public Property Translator As ProgramTranslator

        Public Property CurrentRunner As ProgramRunner

        Public Property PausedRunner As ProgramRunner

        Public Runners As Dictionary(Of Parser, ProgramRunner)
        Public SubRunners As Dictionary(Of Parser, ProgramRunner)
        Public Evaluating As Boolean

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
                        If runner.BreakMode Then
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
                _PausedRunner.BreakMode = False
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

        Public ReadOnly Property BreakMode As Boolean
            Get
                For Each runner In Runners.Values
                    If runner.BreakMode Then Return True
                Next

                For Each runner In SubRunners.Values
                    If runner.BreakMode Then Return True
                Next

                Return False
            End Get
        End Property


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
                    MainRunner.BreakMode = False
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

        Public Sub Disopese()
            Me.Parsers = Nothing
            Me.CurrentRunner = Nothing
            Me.GlobalParser = Nothing
            Me.Runners.Clear()
            Me.SubRunners.Clear()
        End Sub

        Public Function GetRunner(file As String) As ProgramRunner
            For Each r In Runners.Values
                If r.Parser.DocPath.ToLower = file.ToLower() Then
                    Dim runner As New ProgramRunner(Me, r.Parser, True)
                    runner.Fields = New Dictionary(Of String, Primitive)(r.Fields)
                    Return runner
                End If
            Next
            Return Nothing
        End Function

        Public Sub UpdateBreakpoints(breakpoints As List(Of Integer), file As String)
            For Each runner In Runners.Values
                If runner.Parser.DocPath.ToLower = file.ToLower() Then
                    runner.Breakpoints = breakpoints
                End If
            Next

            For Each runner In SubRunners.Values
                If runner.Parser.DocPath.ToLower = file.ToLower() Then
                    runner.Breakpoints = breakpoints
                End If
            Next
        End Sub
    End Class

End Namespace

