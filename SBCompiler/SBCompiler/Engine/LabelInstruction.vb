Imports System

Namespace Microsoft.SmallVisualBasic.Engine
    <Serializable>
    Public Class LabelInstruction
        Inherits Instruction

        Public Property LabelName As String

        Public Overrides ReadOnly Property InstructionType As InstructionType
            Get
                Return InstructionType.LabelInstruction
            End Get
        End Property
    End Class
End Namespace
