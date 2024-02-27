Imports System

Namespace Microsoft.SmallVisualBasic.Engine
    <Serializable>
    Public MustInherit Class Instruction
        Public Property LineNumber As Integer
        Public MustOverride ReadOnly Property InstructionType As InstructionType

        Public MustOverride Function Execute(runner As ProgramRunner) As String
    End Class
End Namespace
