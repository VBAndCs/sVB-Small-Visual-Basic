Imports System
Imports Microsoft.SmallBasic.Expressions

Namespace Microsoft.SmallBasic.Engine
    <Serializable>
    Public Class IfNotGotoInstruction
        Inherits Instruction

        Public Property Condition As Expression
        Public Property LabelName As String

        Public Overrides ReadOnly Property InstructionType As InstructionType
            Get
                Return InstructionType.IfNotGoto
            End Get
        End Property
    End Class
End Namespace
