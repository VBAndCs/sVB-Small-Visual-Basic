﻿Imports System.Reflection.Emit
Imports Microsoft.SmallBasic.Completion

Namespace Microsoft.SmallBasic.Statements

    Public Class JumpLoopStatement
        Inherits Statement

        Public UpLevel As Integer

        Public Overrides Sub AddSymbols(symbolTable As SymbolTable)
            MyBase.AddSymbols(symbolTable)
            Dim parentStatement As Statement = Me
            Do
                parentStatement = parentStatement.Parent
                If parentStatement Is Nothing Then
                    symbolTable.Errors.Add(New [Error](StartToken, "ExitLoop can only appear insde For and While blocks."))
                    Return
                End If

                If TypeOf parentStatement Is ForStatement OrElse
                    TypeOf parentStatement Is WhileStatement Then Return
            Loop

            symbolTable.Errors.Add(New [Error](StartToken, $"{StartToken.Text} can only appear insde For and While blocks."))
        End Sub

        Public Overrides Sub EmitIL(scope As CodeGenScope)
            Dim parentStatement As Statement = Me
            Dim level = UpLevel + 1
            Dim loopStatement As LoopStatement = Nothing

            Do
                If parentStatement.Parent Is Nothing OrElse level = 0 Then
                    If loopStatement IsNot Nothing Then
                        If StartToken.Token = Token.ExitLoop Then
                            scope.ILGenerator.Emit(OpCodes.Br, loopStatement.ExitLabel)
                        Else
                            scope.ILGenerator.Emit(OpCodes.Br, loopStatement.ContinueLabel)
                        End If
                    End If
                    Return
                End If

                parentStatement = parentStatement.Parent

                If TypeOf parentStatement Is ForStatement OrElse
                            TypeOf parentStatement Is WhileStatement Then

                    loopStatement = parentStatement
                    level -= 1
                End If
            Loop
        End Sub

    End Class
End Namespace