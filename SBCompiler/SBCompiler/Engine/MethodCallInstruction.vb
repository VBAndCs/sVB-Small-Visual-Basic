Imports System
Imports Microsoft.SmallBasic.Expressions

Namespace Microsoft.SmallBasic.Engine
    <Serializable>
    Public Class MethodCallInstruction
        Inherits Instruction

        Public Property MethodExpression As MethodCallExpression

        Public Overrides ReadOnly Property InstructionType As InstructionType
            Get
                Return InstructionType.MethodCall
            End Get
        End Property
    End Class
End Namespace
