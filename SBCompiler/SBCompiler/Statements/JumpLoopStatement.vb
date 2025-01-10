﻿Imports System.Reflection.Emit
Imports Microsoft.SmallVisualBasic.Completion
Imports Microsoft.SmallVisualBasic.Engine

Namespace Microsoft.SmallVisualBasic.Statements

    Public Class JumpLoopStatement
        Inherits Statement

        Public UpLevel As Integer

        Public Overrides Function GetStatementAt(lineNumber As Integer) As Statement
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
            Dim level = If(UpLevel = -1, 1000, UpLevel + 1)

            Do
                If parentStatement.Parent Is Nothing OrElse level = 0 Then
                    Return loops
                End If

                parentStatement = parentStatement.Parent

                If TypeOf parentStatement Is ForStatement OrElse
                        TypeOf parentStatement Is ForEachStatement OrElse
                        TypeOf parentStatement Is WhileStatement Then
                    loops.Add(parentStatement)
                    level -= 1
                End If
            Loop
        End Function

        Public Overrides Sub EmitIL(scope As CodeGenScope)
            If scope.ForGlobalHelp Then Return

            Dim loops = GetParentLoops()

            If loops.Count > 0 Then
                If StartToken.Type = TokenType.ExitLoop Then
                    scope.ILGenerator.Emit(OpCodes.Br, loops(loops.Count - 1).ExitLabel)
                Else
                    scope.ILGenerator.Emit(OpCodes.Br, loops(loops.Count - 1).ContinueLabel)
                End If
            End If
        End Sub

        Public Overrides Sub PopulateCompletionItems(
                        bag As CompletionBag,
                        line As Integer,
                        column As Integer,
                        globalScope As Boolean
                   )

            If Me.StartToken.IsBefore(line, column) Then
                bag.CompletionItems.Clear()
                bag.CompletionItems.Add(New CompletionItem() With {
                    .Key = "-",
                    .DisplayName = "-",
                    .ItemType = CompletionItemType.Keyword
                })
            End If
        End Sub

        Public Overrides Function ToString() As String
            Return StartToken.Text & If(UpLevel = -1, " *", New String("-"c, UpLevel)) & vbCrLf
        End Function

        Public Overrides Function Execute(runner As ProgramRunner) As Statement
            Return Me
        End Function

        Public Overrides Function ToVB(symbolTable As SymbolTable) As String
            Dim parentLoops = Me.GetParentLoops()
            Dim loopType = If(parentLoops(0).StartToken.Type = TokenType.While,
                                            "While",
                                            "For")

            If parentLoops.Count = 1 OrElse UpLevel = 0 Then
                If StartToken.Type = TokenType.ExitLoop Then
                    Return "Exit " & loopType
                Else
                    Return "Continue " & loopType
                End If
            Else
                ' ToDo: use GoTo
            End If
        End Function
    End Class
End Namespace