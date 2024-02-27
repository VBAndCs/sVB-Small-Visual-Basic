Imports System
Imports Microsoft.SmallVisualBasic.Expressions

Namespace Microsoft.SmallVisualBasic.Engine
    <Serializable>
    Public Class FieldAssignmentInstruction
        Inherits Instruction

        Public Property FieldName As String
        Public Property RightSide As Expression

        Public Overrides ReadOnly Property InstructionType As InstructionType
            Get
                Return InstructionType.FieldAssignment
            End Get
        End Property

        Public Overrides Function Execute(runner As ProgramRunner) As String
            runner.Fields(_FieldName) = _RightSide.Evaluate(runner)
            Return Nothing
        End Function
    End Class
End Namespace
