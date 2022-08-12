Imports System.Reflection.Emit
Imports Microsoft.SmallBasic.Completion

Namespace Microsoft.SmallBasic.Statements

    Public Class JumpLoopStatement
        Inherits Statement

        Public UpLevel As Integer

        Public Overrides Function GetStatement(lineNumber) As Statement
            If lineNumber = StartToken.Line Then Return Me
            Return Nothing
        End Function


        Public Overrides Sub AddSymbols(symbolTable As SymbolTable)
            MyBase.AddSymbols(symbolTable)
            Dim loops = GetParentLoops()

            If loops.Count = 0 Then
                symbolTable.Errors.Add(New [Error](StartToken, $"{StartToken.Text} can only appear insde For and While blocks."))
            Else
                If UpLevel > loops.Count - 1 Then
                    symbolTable.Errors.Add(New [Error](StartToken, $"There is only {loops.Count} loop{If(loops.Count = 1, "", "s")} but you want To {StartToken.Text.Replace("Loop", "")} {UpLevel + 1} loops"))
                End If

                For Each LoopStatement In loops
                    LoopStatement.JumpLoopStatements.Add(Me)
                Next
            End If

        End Sub

        Function GetParentLoops() As List(Of LoopStatement)
            Dim loops As New List(Of LoopStatement)

            Dim parentStatement As Statement = Me
            Dim level = UpLevel + 1

            Do
                If parentStatement.Parent Is Nothing OrElse level = 0 Then
                    Return loops
                End If

                parentStatement = parentStatement.Parent

                If TypeOf parentStatement Is ForStatement OrElse
                            TypeOf parentStatement Is WhileStatement Then
                    loops.Add(parentStatement)
                    level -= 1
                End If
            Loop
        End Function

        Public Overrides Sub EmitIL(scope As CodeGenScope)
            Dim loops = GetParentLoops()

            If loops.Count > 0 Then
                If StartToken.Token = Token.ExitLoop Then
                    scope.ILGenerator.Emit(OpCodes.Br, loops(loops.Count - 1).ExitLabel)
                Else
                    scope.ILGenerator.Emit(OpCodes.Br, loops(loops.Count - 1).ContinueLabel)
                End If
            End If
        End Sub

    End Class
End Namespace