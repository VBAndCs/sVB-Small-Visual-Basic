Imports System
Imports Microsoft.SmallVisualBasic.Expressions

Namespace Microsoft.SmallVisualBasic.Engine
    <Serializable>
    Public Class IfGotoInstruction
        Inherits Instruction

        Public Property Condition As Expression
        Public Property LabelName As String

        Public Overrides ReadOnly Property InstructionType As InstructionType
            Get
                Return InstructionType.IfGoto
            End Get
        End Property

        Public Overrides Function Execute(runner As ProgramRunner) As String
            If Library.Primitive.ConvertToBoolean(_Condition.Evaluate(runner)) Then
                Return _LabelName
            End If

            Return Nothing
        End Function
    End Class
End Namespace
