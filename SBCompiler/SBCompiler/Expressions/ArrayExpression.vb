Imports System.Reflection
Imports System.Reflection.Emit
Imports Microsoft.SmallVisualBasic.Library
Imports Microsoft.SmallVisualBasic.Statements

Namespace Microsoft.SmallVisualBasic.Expressions
    <Serializable>
    Public Class ArrayExpression
        Inherits Expression

        Public Property LeftHand As Expression
        Public Property Indexer As Expression



        Public Overrides Sub AddSymbols(symbolTable As SymbolTable)
            MyBase.AddSymbols(symbolTable)
            If LeftHand IsNot Nothing Then
                LeftHand.Parent = Me.Parent
                LeftHand.AddSymbols(symbolTable)
            End If

            If Indexer IsNot Nothing Then
                Indexer.Parent = Me.Parent
                Indexer.AddSymbols(symbolTable)
            End If
        End Sub

        Public Sub AddSymbolInitialization(symbolTable As SymbolTable)
            Dim arrExpr = TryCast(LeftHand, ArrayExpression)

            If arrExpr Is Nothing Then
                Dim idExpr = TryCast(LeftHand, IdentifierExpression)
                idExpr?.AddSymbolInitialization(symbolTable)
            Else
                arrExpr.AddSymbolInitialization(symbolTable)
            End If
        End Sub

        Public Overrides Sub EmitIL(scope As CodeGenScope)
            If LeftHand Is Nothing Then Return

            LeftHand.EmitIL(scope)
            If Indexer Is Nothing Then
                ' This is just for intellisense info of the global file
                ' that may contain errors which are ignored at design time.
                ' This will not affect the compilation.
                LiteralExpression.Zero.EmitIL(scope)
            Else
                Indexer.EmitIL(scope)
            End If

            scope.ILGenerator.EmitCall(OpCodes.Call, scope.TypeInfoBag.GetArrayValue, Nothing)
        End Sub

        Public Sub EmitILForSetter(scope As CodeGenScope)
            If LeftHand Is Nothing OrElse Indexer Is Nothing Then Return

            LeftHand.EmitIL(scope)
            Indexer.EmitIL(scope)
            scope.ILGenerator.EmitCall(OpCodes.Call, scope.TypeInfoBag.SetArrayValue, Nothing)
            Dim arrayExpression = TryCast(LeftHand, ArrayExpression)
            Dim identifierExpression = TryCast(LeftHand, IdentifierExpression)

            If arrayExpression IsNot Nothing Then
                arrayExpression.EmitILForSetter(scope)
            ElseIf identifierExpression IsNot Nothing Then
                Dim subroutine = identifierExpression.Subroutine
                Dim identifier = identifierExpression.Identifier

                Dim var = scope.GetLocalBuilder(subroutine, identifier)
                If var IsNot Nothing Then
                    scope.ILGenerator.Emit(OpCodes.Stloc, var)

                ElseIf scope.Fields.ContainsKey(identifier.LCaseText) Then
                    Dim field = scope.Fields(identifierExpression.Identifier.LCaseText)
                    scope.ILGenerator.Emit(OpCodes.Stsfld, field)

                Else
                    scope.SymbolTable.Errors.Add(New [Error](identifier, $"The variable `{identifier.Text}` is used before being initialized."))
                End If

            End If
        End Sub

        Public Overrides Function ToString() As String
            Return $"{LeftHand}[{Indexer}]"
        End Function

        Public Overrides Function InferType(symbolTable As SymbolTable) As VariableType
            If LeftHand.InferType(symbolTable) = VariableType.String Then
                Return VariableType.String
            Else
                Return VariableType.Any
            End If
        End Function

        Public Overrides Function Evaluate(runner As Engine.ProgramRunner) As Primitive
            Dim idEpr = TryCast(_LeftHand, IdentifierExpression)
            Dim value As Primitive = Nothing

            If idEpr IsNot Nothing Then
                Dim fields = runner.Fields
                If Not fields.TryGetValue(runner.GetKey(idEpr.Identifier), value) Then
                    value = ""
                End If

                Return Primitive.GetArrayValue(value, _Indexer.Evaluate(runner))
            End If

            Dim arrExpr = TryCast(_LeftHand, ArrayExpression)

            If arrExpr IsNot Nothing Then
                Return Primitive.GetArrayValue(arrExpr.Evaluate(runner), _Indexer.Evaluate(runner))
            End If

            Return Nothing
        End Function

    End Class
End Namespace
