Imports System
Imports Microsoft.SmallVisualBasic.Expressions

Namespace Microsoft.SmallVisualBasic.Engine
    <Serializable>
    Public Class MethodCallInstruction
        Inherits Instruction

        Public Property MethodExpression As MethodCallExpression

        Public Overrides ReadOnly Property InstructionType As InstructionType
            Get
                Return InstructionType.MethodCall
            End Get
        End Property

        Public Overrides Function Execute(runner As ProgramRunner) As String
            _MethodExpression.Evaluate(runner)
            Return Nothing
        End Function
    End Class
End Namespace
