Imports System
Imports System.Reflection
Imports Microsoft.SmallVisualBasic.Expressions

Namespace Microsoft.SmallVisualBasic.Engine
    <Serializable>
    Public Class PropertyAssignmentInstruction
        Inherits Instruction

        Public Property PropertyInfo As PropertyInfo
        Public Property RightSide As Expression

        Public Overrides ReadOnly Property InstructionType As InstructionType
            Get
                Return InstructionType.PropertyAssignment
            End Get
        End Property

        Public Overrides Function Execute(runner As ProgramRunner) As String
            _PropertyInfo.SetValue(Nothing, _RightSide.Evaluate(runner), Nothing)
            Return Nothing
        End Function
    End Class
End Namespace
