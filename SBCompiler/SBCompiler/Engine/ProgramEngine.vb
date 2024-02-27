Imports System
Imports System.Collections.Generic
Imports System.Reflection
Imports System.Threading
Imports Microsoft.SmallVisualBasic.Expressions
Imports Microsoft.SmallVisualBasic.Library

Namespace Microsoft.SmallVisualBasic.Engine
    Public Class ProgramEngine
        Inherits MarshalByRefObject

        Public ReadOnly Property Breakpoints As List(Of Integer)

        Public Property DebuggerAppDomain As AppDomain

        Public Property Compiler As Compiler

        Public Property CurrentDebuggerState As DebuggerState

        Public Property CurrentInstruction As Instruction

        Public Property Translator As ProgramTranslator

        Public Property Runner As ProgramRunner

        Public Event DebuggerStateChanged As EventHandler
        Public Event EngineUnloaded As EventHandler
        Public Event LineNumberChanged As EventHandler

        Public Sub New(compiler As Compiler)
            Me.Compiler = compiler
            '_Translator = New ProgramTranslator(Me.Compiler)
            '_Translator.TranslateProgram()
            Dim info As New AppDomainSetup() With {
                .LoaderOptimization = LoaderOptimization.MultiDomain
            }
            _DebuggerAppDomain = AppDomain.CreateDomain("Debuggee", Nothing, info)
            _DebuggerAppDomain.SetData("ProgramEngine", Me)
            _DebuggerAppDomain.DoCallBack(AddressOf InitializeRunner)
            AddHandler _DebuggerAppDomain.DomainUnload, Sub() RaiseEvent EngineUnloaded(Me, EventArgs.Empty)
        End Sub

        Private Sub InitializeRunner()
            _Runner = New ProgramRunner(_DebuggerAppDomain)
            AddHandler _Runner.DebuggerStateChanged, Sub() RaiseDebuggerStateChangedEvent(Runner.DebuggerState)
            AddHandler _Runner.LineNumberChanged, Sub() RaiseLineNumberChangedEvent(Runner.CurrentInstruction)
        End Sub

        Public Sub AddBreakpoint(lineNumber As Integer)
            DebuggerAppDomain.DoCallBack(Sub() Runner.Breakpoints.Add(lineNumber))
        End Sub

        Public Sub [Continue]()
            DebuggerAppDomain.DoCallBack(Sub() Runner.Continue())
        End Sub

        Public Sub Pause()
            DebuggerAppDomain.DoCallBack(Sub() Runner.Pause())
        End Sub

        Private Sub RaiseDebuggerStateChangedEvent(currentState As DebuggerState)
            CurrentDebuggerState = currentState
            RaiseEvent DebuggerStateChanged(Me, EventArgs.Empty)
        End Sub

        Private Sub RaiseLineNumberChangedEvent(currentInstruction As Instruction)
            Me.CurrentInstruction = currentInstruction
            RaiseEvent LineNumberChanged(Me, EventArgs.Empty)
        End Sub

        Public Sub StepInto()
            DebuggerAppDomain.DoCallBack(Sub() Runner.StepInto())
        End Sub

        Public Sub StepOver()
            DebuggerAppDomain.DoCallBack(Sub() Runner.StepOver())
        End Sub

        Public Sub Reset()
            DebuggerAppDomain.DoCallBack(Sub() Runner.Reset())
        End Sub

        Public Sub RunProgram()
            _DebuggerAppDomain.DoCallBack(
                Sub() Runner.RunProgram(stopOnFirstInstruction:=False)
            )
        End Sub

        Public Sub RunProgram(stopOnFirstInstruction As Boolean)
            DebuggerAppDomain.DoCallBack(Sub() Runner.RunProgram(stopOnFirstInstruction:=True))
        End Sub
    End Class


End Namespace

