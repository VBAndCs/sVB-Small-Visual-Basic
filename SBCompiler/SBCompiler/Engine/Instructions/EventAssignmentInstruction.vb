Imports System
Imports System.Reflection

Namespace Microsoft.SmallVisualBasic.Engine
    <Serializable>
    Public Class EventAssignmentInstruction
        Inherits Instruction

        Public Property EventInfo As EventInfo
        Public Property SubroutineName As String

        Public Overrides ReadOnly Property InstructionType As InstructionType
            Get
                Return InstructionType.EventAssignment
            End Get
        End Property

        Public Overrides Function Execute(runner As ProgramRunner) As String
            Dim subInstructs = runner.SubroutineInstructions(_SubroutineName)
            '_EventInfo.AddEventHandler(Nothing, Sub() runner.ExecuteInstructions(subInstructs))
            Return Nothing
        End Function
    End Class
End Namespace
