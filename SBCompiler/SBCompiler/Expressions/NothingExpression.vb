Imports System
Imports System.Globalization
Imports System.Reflection.Emit
Imports Microsoft.SmallVisualBasic.Library

Namespace Microsoft.SmallVisualBasic.Expressions
    <Serializable>
    Public Class NothingExpression
        Inherits Expression

        Public Sub New()
            MyBase.New()
        End Sub

        Public Sub New(nothingToken As Token)
            _NothingToken = nothingToken
        End Sub

        Public Property NothingToken As Token

        Public Overrides Sub AddSymbols(symbolTable As SymbolTable)
            MyBase.AddSymbols(symbolTable)
            _NothingToken.Parent = Me.Parent

            Dim st = TryCast(Me.Parent, Statements.AssignmentStatement)

            If st IsNot Nothing Then
                If st.RightValue IsNot Me Then
                    symbolTable.Errors.Add(New [Error](NothingToken, "The keyword Nothing can't be a part of ant other expression. It can only be assigned to an event to remove its handler."))
                    Return
                End If

                Dim propertyExpr = TryCast(st.LeftValue, PropertyExpression)
                If propertyExpr IsNot Nothing AndAlso propertyExpr.IsEvent Then
                    Return
                End If
            End If

            symbolTable.Errors.Add(New [Error](NothingToken, "The keyword Nothing can only be assigned to an event to remove its handler."))

        End Sub

        Public Overrides Sub EmitIL(scope As CodeGenScope)
        End Sub

        Public Overrides Function ToString() As String
            Return NothingToken.Text
        End Function

        Public Overrides Function InferType(symbolTable As SymbolTable) As VariableType
            Return VariableType.Any
        End Function

        Public Overrides Function Evaluate(runner As Engine.ProgramRunner) As Primitive
            Return Nothing
        End Function

        Public Overrides Function ToVB() As String
            Return NothingToken.Text
        End Function
    End Class

End Namespace
