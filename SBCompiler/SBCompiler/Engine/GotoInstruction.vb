Imports System

Namespace Microsoft.SmallBasic.Engine
    <Serializable>
    Public Class GotoInstruction
        Inherits Instruction

        Public Property LabelName As String

        Public Overrides ReadOnly Property InstructionType As InstructionType
            Get
                Return InstructionType.Goto
            End Get
        End Property
    End Class
End Namespace
