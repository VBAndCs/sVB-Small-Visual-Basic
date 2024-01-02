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
    End Class
End Namespace
