Imports System

Namespace Microsoft.SmallVisualBasic.Engine
    <Serializable>
    Public Class SubroutineCallInstruction
        Inherits Instruction

        Public Property SubroutineName As String

        Public Overrides ReadOnly Property InstructionType As InstructionType
            Get
                Return InstructionType.SubroutineCall
            End Get
        End Property

        Public Overrides Function Execute(runner As ProgramRunner) As String
            Dim command = runner.DebuggerCommand

            If command = DebuggerCommand.StepOver Then
                command = DebuggerCommand.Run
            ElseIf command = DebuggerCommand.StepInto Then
                runner.Pause()
            End If

            runner.ExecuteInstructions(runner.SubroutineInstructions(_SubroutineName))
            runner.DebuggerCommand = command
            Return Nothing
        End Function
    End Class
End Namespace
