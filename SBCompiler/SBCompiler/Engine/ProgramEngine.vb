Imports System
Imports System.Collections.Generic
Imports System.Reflection
Imports System.Threading
Imports Microsoft.SmallVisualBasic.Expressions
Imports Microsoft.SmallVisualBasic.Library

Namespace Microsoft.SmallVisualBasic.Engine
    Public Class ProgramEngine

        Public Event DebuggerStateChanged()

        Public Property Parsers As List(Of Parser)

        Public Property CurrentDebuggerState As DebuggerState

        Public ReadOnly Property CurrentLineNumber As Integer
            Get
                Return CurrentRunner.CurrentLineNumber
            End Get
        End Property

        Public Property Translator As ProgramTranslator

        Public Property CurrentRunner As ProgramRunner

        Public Runners As Dictionary(Of Parser, ProgramRunner)

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
            runners = New Dictionary(Of Parser, ProgramRunner)
            _CurrentRunner = New ProgramRunner(Me, mainParser)
            If GlobalParser IsNot Nothing Then
                Dim __ = New ProgramRunner(Me, GlobalParser)
            End If
        End Sub

        Friend Sub RaiseDebuggerStateChangedEvent()
            RaiseEvent DebuggerStateChanged()
        End Sub

        Public Sub AddBreakpoint(lineNumber As Integer)
            CurrentRunner.Breakpoints.Add(lineNumber)
        End Sub

        Public Sub [Continue]()
            CurrentRunner.DebuggerCommand = DebuggerCommand.Run
            CurrentRunner.Continue()
        End Sub

        Public Sub Pause()
            CurrentRunner.Pause()
        End Sub


        Public Sub StepInto()
            CurrentRunner.StepInto()
        End Sub

        Public Sub StepOut()
            CurrentRunner.StepOut()
        End Sub

        Public Sub ShortStepOut()
            CurrentRunner.ShortStepOut()
        End Sub

        Public Sub StepOver()
            CurrentRunner.StepOver()
        End Sub

        Public Sub ShortStepOver()
            CurrentRunner.ShortStepOver()
        End Sub

        Public Sub Reset()
            CurrentRunner.Reset()
        End Sub


        Friend StopOnFirstStaement As Boolean
        Public BreakMode As Boolean

        Public Sub RunProgram(Optional stopOnFirstStaement As Boolean = False)
            Me.StopOnFirstStaement = stopOnFirstStaement
            DoRunProgram()
        End Sub

        Private Sub DoRunProgram()
            Dim debuggerThread As New Thread(
                Sub()
                    WinForms.GeometricPath.CreatePath()
                    Try
                        Library.TextWindow.ClearIfLoaded()
                    Catch
                    End Try

                    Library.Program.IsTerminated = False
                    CurrentRunner.currentThread = Thread.CurrentThread
                    Dim result = CurrentRunner.Execute(CurrentRunner.currentParser.ParseTree)

                    ' Allow event hanlers to break
                    If CurrentRunner.DebuggerCommand = DebuggerCommand.StepOut Then
                        CurrentRunner.DebuggerCommand = DebuggerCommand.Run
                    End If
                    CurrentRunner.StepAround = False
                    BreakMode = False

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
        End Sub


    End Class


End Namespace

