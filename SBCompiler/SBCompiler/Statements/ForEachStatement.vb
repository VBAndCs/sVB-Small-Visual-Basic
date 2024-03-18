Imports System.Reflection.Emit
Imports System.Text
Imports Microsoft.SmallVisualBasic.Completion
Imports Microsoft.SmallVisualBasic.Engine
Imports Microsoft.SmallVisualBasic.Expressions

Namespace Microsoft.SmallVisualBasic.Statements
    Public Class ForEachStatement
        Inherits LoopStatement

        Private Shared loopNo As Integer

        Public ForEachToken As Token
        Public Iterator As Token
        Public InToken As Token
        Public ArrayExpression As Expression
        Public Subroutine As SubroutineStatement

        Public Overrides Function GetStatementAt(lineNumber As Integer) As Statement
            If lineNumber < StartToken.Line Then Return Nothing
            If lineNumber > EndLoopToken.Line Then Return Nothing
            If lineNumber <= Iterator.Line Then Return Me
            If lineNumber <= InToken.Line Then Return Me
            If lineNumber <= ArrayExpression?.EndToken.Line Then Return Me

            For Each statment In Body
                Dim st = statment.GetStatementAt(lineNumber)
                If st IsNot Nothing Then Return st
            Next

            If lineNumber <= EndLoopToken.Line Then Return Me

            Return Nothing
        End Function

        Public Overrides Function GetKeywords() As LegalTokens
            Dim spans As New LegalTokens
            spans.Add(ForEachToken)
            spans.Add(InToken)
            For Each statement In JumpLoopStatements
                spans.Add(statement.StartToken)
            Next
            spans.Add(EndLoopToken)
            Return spans
        End Function

        Dim whileStatement As List(Of Statement)

        Public Overrides Sub AddSymbols(symbolTable As SymbolTable)
            MyBase.AddSymbols(symbolTable)
            Iterator.Parent = Me
            ForEachToken.Parent = Me
            InToken.Parent = Me

            If Iterator.Type <> TokenType.Illegal Then
                Dim id As New IdentifierExpression() With {
                    .Identifier = Iterator,
                    .Subroutine = Subroutine
                }
                Iterator.SymbolType = If(Subroutine Is Nothing,
                    CompletionItemType.GlobalVariable,
                    CompletionItemType.LocalVariable
                )
                symbolTable.AddIdentifier(Iterator)

                Dim key = symbolTable.AddVariable(
                    id,
                    "The ForEach loop iterator",
                    Subroutine IsNot Nothing
                 )

                If key <> "" Then
                    Dim varType = WinForms.PreCompiler.GetVarType(Iterator.Text)
                    If varType <> VariableType.Any Then
                        symbolTable.InferedTypes(key) = varType
                    End If
                End If

                If symbolTable.IsGlobalVar(id) Then
                    id.AddSymbolInitialization(symbolTable)
                End If
            End If

            If ArrayExpression IsNot Nothing Then
                ArrayExpression.Parent = Me
                ArrayExpression.AddSymbols(symbolTable)
            End If

            For Each item In Body
                item.Parent = Me
                item.AddSymbols(symbolTable)
            Next

            whileStatement = LowerToWhile(symbolTable, Me.Subroutine, ForEachToken.Line)

        End Sub

        Private Function LowerToWhile(symbolTable As SymbolTable, Subroutine As Statements.SubroutineStatement, lineOffset As Integer) As List(Of Statement)
            Dim tempRoutine = SubroutineStatement.Current
            SubroutineStatement.Current = Subroutine
            CodeGenerator.IgnoreVarErrors = True

            loopNo += 1
            Dim code =
                $"__foreach__tmparray__{loopNo} = {ArrayExpression}
                    __foreach__indexes__{loopNo} = Array.GetAllIndices(__foreach__tmparray__{loopNo})
                   __foreach__count__{loopNo} = Array.GetItemCount(__foreach__indexes__{loopNo})
                   __foreach__counter__{loopNo} = 1
                   While  __foreach__counter__{loopNo} <= __foreach__count__{loopNo}
                            {Iterator.Text} =__foreach__tmparray__{loopNo}[__foreach__Indexes__{loopNo}[__foreach__counter__{loopNo}]]
                            __foreach__counter__{loopNo} = __foreach__counter__{loopNo} + 1    
                  Wend"

            Dim _parser = Parser.Parse(code, symbolTable, lineOffset)

            CodeGenerator.IgnoreVarErrors = False
            SubroutineStatement.Current = tempRoutine

            CType(_parser.ParseTree.Last, WhileStatement).Body.AddRange(Me.Body)
            Return _parser.ParseTree
        End Function

        Public Overrides Sub PrepareForEmit(scope As CodeGenScope)
            If scope.ForGlobalHelp Then Return

            For Each item In CType(whileStatement.Last, WhileStatement).Body
                item.PrepareForEmit(scope)
            Next
        End Sub

        Public Overrides Property ExitLabel As Label
            Get
                Return CType(whileStatement.Last, WhileStatement).ExitLabel
            End Get
            Set(value As Label)
                CType(whileStatement.Last, WhileStatement).ExitLabel = value
            End Set
        End Property

        Public Overrides Property ContinueLabel As Label
            Get
                Return CType(whileStatement.Last, WhileStatement).ContinueLabel
            End Get
            Set(value As Label)
                CType(whileStatement.Last, WhileStatement).ContinueLabel = value
            End Set
        End Property

        Public Overrides Sub EmitIL(scope As CodeGenScope)
            If scope.ForGlobalHelp Then Return

            For Each statement In whileStatement
                statement.EmitIL(scope)
            Next
        End Sub

        Public Overrides Sub PopulateCompletionItems(
                       bag As CompletionBag,
                       line As Integer,
                       column As Integer,
                       globalScope As Boolean
                   )

            If ForEachToken.Line = line AndAlso column <= ForEachToken.EndColumn Then
                CompletionHelper.FillAllGlobalItems(bag, globalScope)

            ElseIf AfterIterator(line, column) Then
                CompletionHelper.FillLocals(bag, Subroutine?.Name.LCaseText)
                If Subroutine Is Nothing Then
                    If bag.ForHelp AndAlso Not Iterator.IsIllegal Then
                        bag.CompletionItems.Add(New CompletionItem() With {
                            .Key = Iterator.LCaseText,
                            .DisplayName = Iterator.Text,
                            .ItemType = CompletionItemType.GlobalVariable,
                            .DefinitionIdintifier = bag.SymbolTable.GlobalVariables(Iterator.LCaseText)
                         })
                    Else
                        CompletionHelper.FillGlobalVariables(bag)
                    End If
                End If

            ElseIf InForEachHeader(line) Then
                CompletionHelper.FillKeywords(bag, {TokenType.To, TokenType.Step})
                CompletionHelper.FillExpressionItems(bag)
                CompletionHelper.FillSubroutines(bag, True)

            ElseIf Not EndLoopToken.IsIllegal AndAlso line = EndLoopToken.Line Then
                bag.CompletionItems.Clear()

            Else
                Dim statement = GetStatementContaining(Body, line)
                If EndLoopToken.IsIllegal Then CompletionHelper.FillKeywords(bag, {TokenType.Next})
                statement?.PopulateCompletionItems(bag, line, column, globalScope:=False)
            End If
        End Sub

        Private Function AfterIterator(line As Integer, column As Integer) As Boolean
            If InToken.IsIllegal Then
                If Iterator.IsIllegal Then Return line = ForEachToken.Line
                Return line = Iterator.Line
            End If

            If line < InToken.Line Then Return True
            If line > InToken.Line Then Return False
            Return column < InToken.Column
        End Function

        Private Function InForEachHeader(line As Integer) As Boolean
            Dim line2 As Integer

            ' Note that for loop parts can be split over multi lines using the _ symbol,
            ' and some of these parts can be missing while the user is typing code

            If ArrayExpression IsNot Nothing Then
                line2 = ArrayExpression.EndToken.Line
            Else
                line2 = InToken.Line
            End If

            Return line <= line2
        End Function

        Public Overrides Function ToString() As String
            Dim sb As New StringBuilder()
            sb.AppendLine($"{ForEachToken.Text} {Iterator.Text} In {ArrayExpression}")

            For Each st In Body
                sb.Append("   ")
                sb.Append(st.ToString())
            Next

            sb.Append(EndLoopToken.Text)
            Return sb.ToString()
        End Function

        Public Overrides Function Execute(runner As ProgramRunner) As Statement
            Dim key = runner.GetKey(Iterator)
            Dim arr = ArrayExpression.Evaluate(runner)
            Dim keys = Library.Array.GetAllIndices(arr)
            Dim start = New Library.Primitive(1)
            Dim [end] = keys.GetItemCount()
            Dim startLine = ForEachToken.Line
            Dim endLine = EndLoopToken.Line
            Dim stepOut = False

            For i = start To [end]
                If i <> start Then runner.CheckForExecutionBreakAtLine(startLine)
                runner.IncreaseDepthOfShortSteps(stepOut)
                runner.Fields(key) = arr.Items(keys.Items(i))

                Dim result = runner.Execute(Body)
                runner.DecreaseDepthOfShortStepOut(stepOut)

                If TypeOf result Is EndDebugging Then
                    If stepOut Then runner.Depth -= 1
                    Return result
                End If

                If TypeOf result Is JumpLoopStatement Then
                    Dim jumpSt = CType(result, JumpLoopStatement)
                    If jumpSt.StartToken.Type = TokenType.ExitLoop Then
                        If stepOut Then runner.Depth -= 1
                        If jumpSt.UpLevel > 0 Then
                            jumpSt.UpLevel -= 1
                            Return jumpSt
                        Else
                            Return Nothing
                        End If

                    ElseIf jumpSt.UpLevel > 0 Then
                        jumpSt.UpLevel -= 1
                        If stepOut Then runner.Depth -= 1
                        Return jumpSt
                    Else
                        jumpSt.UpLevel = 0
                        Continue For
                    End If

                ElseIf TypeOf result Is ReturnStatement Then
                    If stepOut Then runner.Depth -= 1
                    Return result

                ElseIf TypeOf result Is GotoStatement Then
                    Dim label = CType(result, GotoStatement).Label
                    If label.Line > EndLoopToken.Line OrElse label.Line < ForEachToken.Line Then
                        If stepOut Then runner.Depth -= 1
                        Return result
                    End If
                End If

                runner.CheckForExecutionBreakAtLine(endLine)
                runner.IncreaseDepthOfShortSteps(stepOut)
            Next

            If stepOut Then runner.Depth -= 1
            Return Nothing
        End Function
    End Class
End Namespace
