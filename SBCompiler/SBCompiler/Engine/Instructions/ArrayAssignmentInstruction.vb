Imports System
Imports Microsoft.SmallVisualBasic.Expressions
Imports Microsoft.SmallVisualBasic.Library

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

        Public Overrides Function Execute(runner As ProgramRunner) As String
            'SetArrayValue(runner, _RightSide.Evaluate(runner))
            Return Nothing
        End Function

        Friend Sub SetArrayValue(runner As ProgramRunner, value As Primitive)
            Dim identifierExpression = TryCast(_ArrayExpression.LeftHand, IdentifierExpression)
            Dim fields = runner.Fields

            If identifierExpression IsNot Nothing Then
                Dim arrName = runner.GetKey(identifierExpression.Identifier)
                Dim arr As Primitive = Nothing
                If Not fields.TryGetValue(arrName, arr) Then
                    arr = Nothing
                End If
                fields(arrName) = Primitive.SetArrayValue(value, arr, _ArrayExpression.Indexer.Evaluate(runner))

            Else
                Dim arrExpr = TryCast(_ArrayExpression.LeftHand, ArrayExpression)
                If arrExpr IsNot Nothing Then
                    SetArrayValue(runner, Primitive.SetArrayValue(
                          value,
                          arrExpr.Evaluate(runner),
                          arrExpr.Indexer.Evaluate(runner))
                    )
                End If
            End If
        End Sub

    End Class
End Namespace
