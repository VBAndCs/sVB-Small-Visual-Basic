Imports System.Collections.Generic
Imports System.Reflection
Imports Microsoft.SmallBasic.Expressions
Imports Microsoft.SmallBasic.Statements

Namespace Microsoft.SmallBasic.Engine
    Public Class ProgramTranslator
        Private _Compiler As Microsoft.SmallBasic.Compiler, _ProgramInstructions As System.Collections.Generic.List(Of Microsoft.SmallBasic.Engine.Instruction), _SubroutineInstructions As System.Collections.Generic.Dictionary(Of String, System.Collections.Generic.List(Of Microsoft.SmallBasic.Engine.Instruction))
        Private labelId As Integer = 10

        Public Property Compiler As Compiler
            Get
                Return _Compiler
            End Get
            Private Set(ByVal value As Compiler)
                _Compiler = value
            End Set
        End Property

        Public Property ProgramInstructions As List(Of Instruction)
            Get
                Return _ProgramInstructions
            End Get
            Private Set(ByVal value As List(Of Instruction))
                _ProgramInstructions = value
            End Set
        End Property

        Public Property SubroutineInstructions As Dictionary(Of String, List(Of Instruction))
            Get
                Return _SubroutineInstructions
            End Get
            Private Set(ByVal value As Dictionary(Of String, List(Of Instruction)))
                _SubroutineInstructions = value
            End Set
        End Property

        Public Sub New(ByVal compiler As Compiler)
            Me.Compiler = compiler
        End Sub

        Private Function CreateNewLabel() As String
            labelId += 10
            Return "__label" & labelId
        End Function

        Public Sub TranslateProgram()
            ProgramInstructions = New List(Of Instruction)()
            SubroutineInstructions = New Dictionary(Of String, List(Of Instruction))()
            TranslateStatements(ProgramInstructions, Compiler.Parser.ParseTree)
        End Sub

        Public Sub TranslateStatements(ByVal instructions As List(Of Instruction), ByVal statements As List(Of Statement))
            For Each statement In statements
                TranslateStatement(instructions, statement)
            Next
        End Sub

        Private Sub TranslateStatement(ByVal instructions As List(Of Instruction), ByVal statement As Statement)
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

        Private Sub TranslateAssignmentStatement(ByVal instructions As List(Of Instruction), ByVal statement As AssignmentStatement)
            Dim identifierExpression As IdentifierExpression = TryCast(statement.LeftValue, IdentifierExpression)

            If identifierExpression IsNot Nothing Then
                instructions.Add(New FieldAssignmentInstruction With {
                    .FieldName = identifierExpression.Identifier.NormalizedText,
                    .LineNumber = statement.StartToken.Line,
                    .RightSide = statement.RightValue
                })
                Return
            End If

            Dim propertyExpression As PropertyExpression = TryCast(statement.LeftValue, PropertyExpression)
            Dim value As EventInfo = Nothing

            If propertyExpression IsNot Nothing Then
                Dim typeInfo = Compiler.TypeInfoBag.Types(propertyExpression.TypeName.NormalizedText)
                Dim value2 As PropertyInfo

                If typeInfo.Events.TryGetValue(propertyExpression.PropertyName.NormalizedText, value) Then
                    Dim identifierExpression2 As IdentifierExpression = TryCast(statement.RightValue, IdentifierExpression)

                    If identifierExpression2 IsNot Nothing Then
                        instructions.Add(New EventAssignmentInstruction With {
                            .EventInfo = value,
                            .LineNumber = statement.StartToken.Line,
                            .SubroutineName = identifierExpression2.Identifier.NormalizedText
                        })
                    End If
                ElseIf typeInfo.Properties.TryGetValue(propertyExpression.PropertyName.NormalizedText, value2) Then
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

        Private Sub TranslateForStatement(ByVal instructions As List(Of Instruction), ByVal statement As ForStatement)
            Dim labelName As String = CreateNewLabel()
            Dim labelName2 As String = CreateNewLabel()
            Dim labelName3 As String = CreateNewLabel()
            Dim labelName4 As String = CreateNewLabel()
            instructions.Add(New FieldAssignmentInstruction With {
                .FieldName = statement.Iterator.NormalizedText,
                .RightSide = statement.InitialValue,
                .LineNumber = statement.StartToken.Line
            })
            instructions.Add(New LabelInstruction With {
                .LabelName = labelName,
                .LineNumber = statement.StartToken.Line
            })

            If statement.StepValue IsNot Nothing Then
                instructions.Add(New IfGotoInstruction With {
                    .Condition = New BinaryExpression With {
                        .LeftHandSide = New LiteralExpression With {
                            .Literal = New TokenInfo With {
                                .Token = Token.NumericLiteral,
                                .Text = "0"
                            }
                        },
                        .Operator = New TokenInfo With {
                            .Token = Token.LessThan
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
                        .Operator = New TokenInfo With {
                            .Token = Token.LessThan
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
                    .Operator = New TokenInfo With {
                        .Token = Token.GreaterThan
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
                .FieldName = statement.Iterator.NormalizedText,
                .RightSide = New BinaryExpression With {
                    .LeftHandSide = New IdentifierExpression With {
                        .Identifier = statement.Iterator,
                        .Subroutine = SubroutineStatement.Current
                    },
                    .Operator = New TokenInfo With {
                        .Token = Token.Addition
                    },
                    .RightHandSide = (If(statement.StepValue IsNot Nothing, statement.StepValue, New LiteralExpression With {
                        .Literal = New TokenInfo With {
                            .Token = Token.NumericLiteral,
                            .Text = "1"
                        }
                    }))
                },
                .LineNumber = statement.StartToken.Line
            })
            instructions.Add(New GotoInstruction With {
                .LabelName = labelName,
                .LineNumber = statement.StartToken.Line
            })
            instructions.Add(New LabelInstruction With {
                .LabelName = labelName4,
                .LineNumber = statement.StartToken.Line
            })
        End Sub

        Private Sub TranslateGotoStatement(ByVal instructions As List(Of Instruction), ByVal statement As GotoStatement)
            instructions.Add(New GotoInstruction With {
                .LabelName = statement.Label.NormalizedText,
                .LineNumber = statement.StartToken.Line
            })
        End Sub

        Private Sub TranslateIfStatement(ByVal instructions As List(Of Instruction), ByVal statement As IfStatement)
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

        Private Sub TranslateLabelStatement(ByVal instructions As List(Of Instruction), ByVal statement As LabelStatement)
            instructions.Add(New LabelInstruction With {
                .LineNumber = statement.StartToken.Line,
                .LabelName = statement.LabelToken.NormalizedText
            })
        End Sub

        Private Sub TranslateMethodCallStatement(ByVal instructions As List(Of Instruction), ByVal statement As MethodCallStatement)
            instructions.Add(New MethodCallInstruction With {
                .LineNumber = statement.StartToken.Line,
                .MethodExpression = statement.MethodCallExpression
            })
        End Sub

        Private Sub TranslateSubroutineStatement(ByVal statement As SubroutineStatement)
            Dim list As List(Of Instruction) = New List(Of Instruction)()
            SubroutineInstructions(statement.Name.NormalizedText) = list
            TranslateStatements(list, statement.Body)
        End Sub

        Private Sub TranslateSubroutineCallStatement(ByVal instructions As List(Of Instruction), ByVal statement As SubroutineCallStatement)
            instructions.Add(New SubroutineCallInstruction With {
                .LineNumber = statement.StartToken.Line,
                .SubroutineName = statement.Name.NormalizedText
            })
        End Sub

        Private Sub TranslateWhileStatement(ByVal instructions As List(Of Instruction), ByVal statement As WhileStatement)
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
