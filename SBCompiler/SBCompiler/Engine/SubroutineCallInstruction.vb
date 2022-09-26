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
    End Class
End Namespace
