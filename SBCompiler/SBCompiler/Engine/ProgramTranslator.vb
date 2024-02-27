Imports System.Collections.Generic
Imports System.Reflection
Imports Microsoft.SmallVisualBasic.Expressions
Imports Microsoft.SmallVisualBasic.Statements

Namespace Microsoft.SmallVisualBasic.Engine
    Public Class ProgramTranslator
        Private labelId As Integer = 10

        Public Property Compiler As Compiler

        Public Property ProgramInstructions As List(Of Instruction)

        Public Property SubroutineInstructions As Dictionary(Of String, List(Of Instruction))

        Dim symbolTable As SymbolTable

        Public Sub New(compiler As Compiler)
            Me.Compiler = compiler
            symbolTable = compiler.Parser.SymbolTable
        End Sub

        Private Function CreateNewLabel() As String
            labelId += 10
            Return "__label" & labelId
        End Function

        Public Sub TranslateProgram()
            _ProgramInstructions = New List(Of Instruction)()
            _SubroutineInstructions = New Dictionary(Of String, List(Of Instruction))()
            TranslateStatements(_ProgramInstructions, _Compiler.Parser.ParseTree)
        End Sub

        Public Sub TranslateStatements(instructions As List(Of Instruction), statements As List(Of Statement))
            For Each statement In statements
                TranslateStatement(instructions, statement)
            Next
        End Sub

        Private Sub TranslateStatement(instructions As List(Of Instruction), statement As Statement)
            If TypeOf statement Is AssignmentStatement Then
                TranslateAssignmentStatement(instructions, TryCast(statement, AssignmentStatement))
            ElseIf TypeOf statement Is ForStatement Then
                TranslateForStatement(instructions, TryCast(statement, ForStatement))
            ElseIf TypeOf statement Is GotoStatement Then
                TranslateGotoStatement(instructions, TryCast(statement, GotoStatement))
            ElseIf TypeOf statement Is IfStatement Then
                TranslateIfStatement(instructions, TryCast(statement, IfStatement))
            ElseIf TypeOf statement Is LabelStatement Then
                TranslateLabelStatement(instructions, TryCast(statement, LabelStatement))
            ElseIf TypeOf statement Is MethodCallStatement Then
                TranslateMethodCallStatement(instructions, TryCast(statement, MethodCallStatement))
            ElseIf TypeOf statement Is SubroutineCallStatement Then
                TranslateSubroutineCallStatement(instructions, TryCast(statement, SubroutineCallStatement))
            ElseIf TypeOf statement Is SubroutineStatement Then
                TranslateSubroutineStatement(TryCast(statement, SubroutineStatement))
            ElseIf TypeOf statement Is WhileStatement Then
                TranslateWhileStatement(instructions, TryCast(statement, WhileStatement))
            End If
        End Sub

        Private Sub TranslateAssignmentStatement(instructions As List(Of Instruction), statement As AssignmentStatement)
            Dim idExpr As IdentifierExpression = TryCast(statement.LeftValue, IdentifierExpression)

            If idExpr IsNot Nothing Then
                instructions.Add(New FieldAssignmentInstruction With {
                    .FieldName = SymbolTable.GetKey(idExpr.Identifier),
                    .LineNumber = statement.StartToken.Line,
                    .RightSide = statement.RightValue
                })
                Return
            End If

            Dim propertyExpression As PropertyExpression = TryCast(statement.LeftValue, PropertyExpression)
            Dim value As EventInfo = Nothing

            If propertyExpression IsNot Nothing Then
                Dim typeInfo = Compiler.TypeInfoBag.Types(propertyExpression.TypeName.LCaseText)
                Dim value2 As PropertyInfo

                If typeInfo.Events.TryGetValue(propertyExpression.PropertyName.LCaseText, value) Then
                    Dim identifierExpression2 As IdentifierExpression = TryCast(statement.RightValue, IdentifierExpression)

                    If identifierExpression2 IsNot Nothing Then
                        instructions.Add(New EventAssignmentInstruction With {
                            .EventInfo = value,
                            .LineNumber = statement.StartToken.Line,
                            .SubroutineName = identifierExpression2.Identifier.LCaseText
                        })
                    End If
                ElseIf typeInfo.Properties.TryGetValue(propertyExpression.PropertyName.LCaseText, value2) Then
                    instructions.Add(New PropertyAssignmentInstruction With {
                        .PropertyInfo = value2,
                        .RightSide = statement.RightValue,
                        .LineNumber = statement.StartToken.Line
                    })
                End If
            Else
                Dim arrayExpression As ArrayExpression = TryCast(statement.LeftValue, ArrayExpression)

                If arrayExpression IsNot Nothing Then
                    instructions.Add(New ArrayAssignmentInstruction With {
                        .ArrayExpression = arrayExpression,
                        .RightSide = statement.RightValue,
                        .LineNumber = statement.StartToken.Line
                    })
                End If
            End If
        End Sub

        Private Sub TranslateForStatement(instructions As List(Of Instruction), statement As ForStatement)
            Dim labelNext As String = CreateNewLabel()
            Dim labelName2 As String = CreateNewLabel()
            Dim labelName3 As String = CreateNewLabel()
            Dim labelName4 As String = CreateNewLabel()
            instructions.Add(New FieldAssignmentInstruction With {
                .FieldName = symbolTable.GetKey(statement.Iterator),
                .RightSide = statement.InitialValue,
                .LineNumber = statement.StartToken.Line
            })
            instructions.Add(New LabelInstruction With {
                .LabelName = labelNext,
                .LineNumber = statement.StartToken.Line
            })

            If statement.StepValue IsNot Nothing Then
                instructions.Add(New IfGotoInstruction With {
                    .Condition = New BinaryExpression With {
                        .LeftHandSide = New LiteralExpression With {
                            .Literal = New Token With {
                                .Type = TokenType.NumericLiteral,
                                .Text = "0"
                            }
                        },
                        .Operator = New Token With {
                            .Type = TokenType.LessThan
                        },
                        .RightHandSide = statement.StepValue
                    },
                    .LabelName = labelName2,
                    .LineNumber = statement.StartToken.Line
                })
                instructions.Add(New IfGotoInstruction With {
                    .Condition = New BinaryExpression With {
                        .LeftHandSide = New IdentifierExpression With {
                            .Identifier = statement.Iterator,
                            .Subroutine = SubroutineStatement.Current
                        },
                        .Operator = New Token With {
                            .Type = TokenType.LessThan
                        },
                        .RightHandSide = statement.FinalValue
                    },
                    .LabelName = labelName4,
                    .LineNumber = statement.StartToken.Line
                })
                instructions.Add(New GotoInstruction With {
                    .LabelName = labelName3,
                    .LineNumber = statement.StartToken.Line
                })
            End If

            instructions.Add(New LabelInstruction With {
                .LabelName = labelName2,
                .LineNumber = statement.StartToken.Line
            })
            instructions.Add(New IfGotoInstruction With {
                .Condition = New BinaryExpression With {
                    .LeftHandSide = New IdentifierExpression With {
                        .Identifier = statement.Iterator,
                        .Subroutine = SubroutineStatement.Current
                    },
                    .Operator = New Token With {
                        .Type = TokenType.GreaterThan
                    },
                    .RightHandSide = statement.FinalValue
                },
                .LabelName = labelName4,
                .LineNumber = statement.StartToken.Line
            })
            instructions.Add(New LabelInstruction With {
                .LabelName = labelName3,
                .LineNumber = statement.StartToken.Line
            })
            TranslateStatements(instructions, statement.Body)
            instructions.Add(New FieldAssignmentInstruction With {
                .FieldName = symbolTable.GetKey(statement.Iterator),
                .RightSide = New BinaryExpression With {
                    .LeftHandSide = New IdentifierExpression With {
                        .Identifier = statement.Iterator,
                        .Subroutine = SubroutineStatement.Current
                    },
                    .Operator = New Token With {
                        .Type = TokenType.Addition
                    },
                    .RightHandSide = (If(statement.StepValue IsNot Nothing, statement.StepValue, New LiteralExpression With {
                        .Literal = New Token With {
                            .Type = TokenType.NumericLiteral,
                            .Text = "1"
                        }
                    }))
                },
                .LineNumber = statement.StartToken.Line
            })
            instructions.Add(New GotoInstruction With {
                .LabelName = labelNext,
                .LineNumber = statement.StartToken.Line
            })
            instructions.Add(New LabelInstruction With {
                .LabelName = labelName4,
                .LineNumber = statement.StartToken.Line
            })
        End Sub

        Private Sub TranslateGotoStatement(instructions As List(Of Instruction), statement As GotoStatement)
            instructions.Add(New GotoInstruction With {
                .LabelName = statement.Label.LCaseText,
                .LineNumber = statement.StartToken.Line
            })
        End Sub

        Private Sub TranslateIfStatement(
                    instructions As List(Of Instruction),
                    statement As IfStatement)

            Dim labelName As String = CreateNewLabel()
            Dim labelName2 As String = CreateNewLabel()
            instructions.Add(New IfNotGotoInstruction With {
                .Condition = statement.Condition,
                .LabelName = labelName2,
                .LineNumber = statement.StartToken.Line
            })
            TranslateStatements(instructions, statement.ThenStatements)
            instructions.Add(New GotoInstruction With {
                .LabelName = labelName,
                .LineNumber = statement.StartToken.Line
            })

            For Each elseIfStatement In statement.ElseIfStatements
                instructions.Add(New LabelInstruction With {
                    .LabelName = labelName2,
                    .LineNumber = statement.StartToken.Line
                })

                labelName2 = CreateNewLabel()
                instructions.Add(New IfNotGotoInstruction With {
                    .Condition = elseIfStatement.Condition,
                    .LabelName = labelName2,
                    .LineNumber = elseIfStatement.StartToken.Line
                })

                TranslateStatements(instructions, elseIfStatement.ThenStatements)
                instructions.Add(New GotoInstruction With {
                    .LabelName = labelName,
                    .LineNumber = statement.StartToken.Line
                })
            Next

            instructions.Add(New LabelInstruction With {
                .LabelName = labelName2,
                .LineNumber = statement.StartToken.Line
            })
            TranslateStatements(instructions, statement.ElseStatements)
            instructions.Add(New LabelInstruction With {
                .LabelName = labelName,
                .LineNumber = statement.StartToken.Line
            })
        End Sub

        Private Sub TranslateLabelStatement(instructions As List(Of Instruction), statement As LabelStatement)
            instructions.Add(New LabelInstruction With {
                .LineNumber = statement.StartToken.Line,
                .LabelName = statement.LabelToken.LCaseText
            })
        End Sub

        Private Sub TranslateMethodCallStatement(instructions As List(Of Instruction), statement As MethodCallStatement)
            instructions.Add(New MethodCallInstruction With {
                .LineNumber = statement.StartToken.Line,
                .MethodExpression = statement.MethodCallExpression
            })
        End Sub

        Private Sub TranslateSubroutineStatement(statement As SubroutineStatement)
            Dim list As List(Of Instruction) = New List(Of Instruction)()
            _SubroutineInstructions(statement.Name.LCaseText) = list
            TranslateStatements(list, statement.Body)
        End Sub

        Private Sub TranslateSubroutineCallStatement(instructions As List(Of Instruction), statement As SubroutineCallStatement)
            instructions.Add(New SubroutineCallInstruction With {
                .LineNumber = statement.StartToken.Line,
                .SubroutineName = statement.Name.LCaseText
            })
        End Sub

        Private Sub TranslateWhileStatement(instructions As List(Of Instruction), statement As WhileStatement)
            Dim labelName As String = CreateNewLabel()
            Dim labelName2 As String = CreateNewLabel()
            instructions.Add(New LabelInstruction With {
                .LabelName = labelName2,
                .LineNumber = statement.StartToken.Line
            })
            instructions.Add(New IfNotGotoInstruction With {
                .Condition = statement.Condition,
                .LabelName = labelName,
                .LineNumber = statement.StartToken.Line
            })
            TranslateStatements(instructions, statement.Body)
            instructions.Add(New GotoInstruction With {
                .LabelName = labelName2,
                .LineNumber = statement.StartToken.Line
            })
            instructions.Add(New LabelInstruction With {
                .LabelName = labelName,
                .LineNumber = statement.StartToken.Line
            })
        End Sub
    End Class
End Namespace
