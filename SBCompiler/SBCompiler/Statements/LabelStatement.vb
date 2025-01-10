﻿Imports System.Globalization
Imports System.Reflection.Emit
Imports Microsoft.SmallVisualBasic.Engine

Namespace Microsoft.SmallVisualBasic.Statements
    Public Class LabelStatement
        Inherits Statement

        Public LabelToken As Token
        Public ColonToken As Token
        Public subroutine As SubroutineStatement

        Public Overrides Function GetStatementAt(lineNumber As Integer) As Statement
            If lineNumber < StartToken.Line Then Return Nothing
            If lineNumber <= LabelToken.Line Then Return Me
            Return Nothing
        End Function

        Public Overrides Sub AddSymbols(symbolTable As SymbolTable)
            MyBase.AddSymbols(symbolTable)
            LabelToken.Parent = Me
            ColonToken.Parent = Me

            If Not LabelToken.IsIllegal Then
                symbolTable.AddLabelDefinition(LabelToken)
                LabelToken.SymbolType = Completion.CompletionItemType.Label
                symbolTable.AddIdentifier(LabelToken)
            End If
        End Sub

        Public Overrides Sub PrepareForEmit(scope As CodeGenScope)
            If scope.ForGlobalHelp Then Return

            Dim value As Label = scope.ILGenerator.DefineLabel()
            scope.Labels.Add(LabelToken.LCaseText, value)
        End Sub

        Public Overrides Sub EmitIL(scope As CodeGenScope)
            If scope.ForGlobalHelp Then Return

            Dim loc = scope.Labels(LabelToken.LCaseText)
            scope.ILGenerator.MarkLabel(loc)
        End Sub

        Public Overrides Function ToString() As String
            Return $"{LabelToken.Text}:" & vbCrLf
        End Function


        Public Overrides Sub PopulateCompletionItems(
                            bag As Completion.CompletionBag,
                            line As Integer,
                            column As Integer,
                            globalScope As Boolean
                    )

            If LabelToken.Contains(line, column, True) Then
                bag.CompletionItems.Add(
                        New Completion.CompletionItem() With {
                                 .DisplayName = LabelToken.Text,
                                 .Key = LabelToken.Text,
                                 .ItemType = Completion.CompletionItemType.Label,
                                 .DefinitionIdintifier = LabelToken
                })

            End If

        End Sub

        Public Overrides Function Execute(runner As ProgramRunner) As statement
            Return Nothing
        End Function

        Public Overrides Function ToVB(symbolTable As SymbolTable) As String
            Return $"{LabelToken.Text}:" & vbCrLf
        End Function
    End Class
End Namespace
