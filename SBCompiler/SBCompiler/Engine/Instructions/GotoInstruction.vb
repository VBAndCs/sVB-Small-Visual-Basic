Imports System

Namespace Microsoft.SmallVisualBasic.Engine
    <Serializable>
    Public Class GotoInstruction
        Inherits Instruction

        Public Property LabelName As String

        Public Overrides ReadOnly Property InstructionType As InstructionType
            Get
                Return InstructionType.Goto
            End Get
        End Property

        Public Overrides Function Execute(runner As ProgramRunner) As String
            Return _LabelName
        End Function
    End Class
End Namespace
