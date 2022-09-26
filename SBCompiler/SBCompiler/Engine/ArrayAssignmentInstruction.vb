Imports System
Imports Microsoft.SmallVisualBasic.Expressions

Namespace Microsoft.SmallVisualBasic.Engine
    <Serializable>
    Public Class ArrayAssignmentInstruction
        Inherits Instruction

        Public Property ArrayExpression As ArrayExpression
        Public Property RightSide As Expression

        Public Overrides ReadOnly Property InstructionType As InstructionType
            Get
                Return InstructionType.ArrayAssignment
            End Get
        End Property
    End Class
End Namespace
