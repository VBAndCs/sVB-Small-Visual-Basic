Imports System.IO
Imports Microsoft.SmallVisualBasic.Expressions
Imports Microsoft.SmallVisualBasic.Library
Imports Microsoft.SmallVisualBasic.Statements
Imports System.Runtime.InteropServices
Imports System.Globalization
Imports Microsoft.SmallBasic

Namespace Microsoft.SmallVisualBasic
    Public Class Parser
        Private codeLines As List(Of String)
        Private _currentLine As Integer
        Private _currentLineEnum As TokenEnumerator
        Private _rewindRequested As Boolean
        Private _FunctionsCall As New List(Of MethodCallExpression)
        Private keepSymbols As Boolean
        Private lineOffset As Integer

        Public ClassName As String = "_SmallVisualBasic_Module1"
        Public IsMainForm As Boolean
        Public IsGlobal As Boolean
        Public FormNames As List(Of String)

        Public Property Errors As List(Of [Error])

        Public ReadOnly Property ParseTree As List(Of Statement)

        Public ReadOnly Property SymbolTable As SymbolTable

        Private Function ConstructWhileStatement(tokenEnum As TokenEnumerator) As WhileStatement
            Dim whileStatement As New WhileStatement With {
                .StartToken = tokenEnum.Current,
                .WhileToken = tokenEnum.Current
            }
            tokenEnum.MoveNext()

            If EatLogicalExpression(tokenEnum, whileStatement.Condition) Then
                ExpectEndOfLine(tokenEnum)
            End If

            tokenEnum = ReadNextLine()
            Dim flag = False

            While tokenEnum IsNot Nothing

                If tokenEnum.Current.Type = TokenType.EndWhile OrElse tokenEnum.Current.Type = TokenType.Wend Then
                    whileStatement.EndLoopToken = tokenEnum.Current
                    flag = True
                    Exit While
                End If

                If tokenEnum.Current.Type = TokenType.Sub OrElse tokenEnum.Current.Type = TokenType.EndSub OrElse
                            tokenEnum.Current.Type = TokenType.Function OrElse tokenEnum.Current.Type = TokenType.EndFunction Then
                    Exit While
                End If

                whileStatement.Body.Add(GetStatementFromTokens(tokenEnum))
                tokenEnum = ReadNextLine()
            End While

            If flag Then
                tokenEnum.MoveNext()
                ExpectEndOfLine(tokenEnum)
            Else
                AddError(ResourceHelper.GetString("EndWhileExpected"))
            End If

            Return whileStatement
        End Function

        Dim commaLine As Integer = -1
        Dim commaColumn As Integer = -1
        Dim callerInfo As CallerInfo
        Private Shared callerInfoParser As New Parser()

        Public Shared Function GetCommaCallerInfo(
                        source As String,
                        line As Integer,
                        column As Integer
                  ) As CallerInfo

            With callerInfoParser
                .commaLine = line
                .commaColumn = column
                .callerInfo = Nothing

                .Parse(New StringReader(source), True)

                .commaLine = -1
                .commaColumn = -1
                Return .callerInfo
            End With
        End Function

        Private Function ConstructForStatement(tokenEnum As TokenEnumerator) As ForStatement
            Dim forStatement As New ForStatement() With {
                .StartToken = tokenEnum.Current,
                .ForToken = tokenEnum.Current,
                .Subroutine = SubroutineStatement.Current
            }

            tokenEnum.MoveNext()
            If EatSimpleIdentifier(tokenEnum, forStatement.Iterator) AndAlso
                    EatToken(tokenEnum, TokenType.Equals, forStatement.EqualsToken) AndAlso
                    EatExpression(tokenEnum, forStatement.InitialValue) AndAlso
                    EatToken(tokenEnum, TokenType.To, forStatement.ToToken) AndAlso
                    EatExpression(tokenEnum, forStatement.FinalValue) AndAlso
                    (Not EatOptionalToken(tokenEnum, TokenType.Step, forStatement.StepToken) OrElse
                        EatExpression(tokenEnum, forStatement.StepValue)) Then

                If commaLine = -2 Then Return Nothing
                ExpectEndOfLine(tokenEnum)
            End If

            If commaLine = -2 Then Return Nothing

            tokenEnum = ReadNextLine()
            Dim flag = False

            While tokenEnum IsNot Nothing

                If tokenEnum.Current.Type = TokenType.EndFor OrElse tokenEnum.Current.Type = TokenType.Next Then
                    forStatement.EndLoopToken = tokenEnum.Current
                    flag = True
                    Exit While
                End If

                If tokenEnum.Current.Type = TokenType.Sub OrElse tokenEnum.Current.Type = TokenType.EndSub OrElse
                        tokenEnum.Current.Type = TokenType.Function OrElse tokenEnum.Current.Type = TokenType.EndFunction Then
                    Exit While
                End If

                forStatement.Body.Add(GetStatementFromTokens(tokenEnum))
                tokenEnum = ReadNextLine()
            End While

            If flag Then
                tokenEnum.MoveNext()
                ExpectEndOfLine(tokenEnum)
            Else
                AddError(ResourceHelper.GetString("EndForExpected"))
            End If

            Return forStatement
        End Function

        Private Function ConstructForEachStatement(tokenEnum As TokenEnumerator) As ForEachStatement
            Dim forEach As New ForEachStatement() With {
                .StartToken = tokenEnum.Current,
                .ForEachToken = tokenEnum.Current,
                .Subroutine = SubroutineStatement.Current
            }

            tokenEnum.MoveNext()
            If EatSimpleIdentifier(tokenEnum, forEach.Iterator) AndAlso
                    EatToken(tokenEnum, TokenType.In, forEach.InToken) AndAlso
                    EatExpression(tokenEnum, forEach.ArrayExpression) Then

                If commaLine = -2 Then Return Nothing
                ExpectEndOfLine(tokenEnum)
            End If

            If commaLine = -2 Then Return Nothing

            tokenEnum = ReadNextLine()
            Dim nextFound = False

            While tokenEnum IsNot Nothing
                If tokenEnum.Current.Type = TokenType.Next Then
                    forEach.EndLoopToken = tokenEnum.Current
                    nextFound = True
                    Exit While
                End If

                If tokenEnum.Current.Type = TokenType.Sub OrElse tokenEnum.Current.Type = TokenType.EndSub OrElse
                        tokenEnum.Current.Type = TokenType.Function OrElse tokenEnum.Current.Type = TokenType.EndFunction Then
                    Exit While
                End If

                forEach.Body.Add(GetStatementFromTokens(tokenEnum))
                tokenEnum = ReadNextLine()
            End While

            If nextFound Then
                tokenEnum.MoveNext()
                ExpectEndOfLine(tokenEnum)
            Else
                AddError(ResourceHelper.GetString("EndForExpected"))
            End If

            Return forEach
        End Function

        Private Function ConstructIfStatement(tokenEnum As TokenEnumerator) As IfStatement
            Dim ifStatement As New IfStatement() With {
                    .StartToken = tokenEnum.Current,
                    .IfToken = tokenEnum.Current
            }

            tokenEnum.MoveNext()

            If EatLogicalExpression(tokenEnum, ifStatement.Condition) AndAlso
                        EatToken(tokenEnum, TokenType.Then, ifStatement.ThenToken) Then
                ExpectEndOfLine(tokenEnum)
            End If

            tokenEnum = ReadNextLine()
            Dim foundEndIf = False
            Dim foundElseIf = False
            Dim foundElse = False

            While tokenEnum IsNot Nothing

                If tokenEnum.Current.Type = TokenType.EndIf Then
                    ifStatement.EndIfToken = tokenEnum.Current
                    foundEndIf = True
                    Exit While
                End If

                If tokenEnum.Current.Type = TokenType.Else Then
                    ifStatement.ElseToken = tokenEnum.Current
                    foundElse = True
                    Exit While
                End If

                If tokenEnum.Current.Type = TokenType.ElseIf Then
                    foundElseIf = True
                    Exit While
                End If

                If tokenEnum.Current.Type = TokenType.Sub OrElse
                            tokenEnum.Current.Type = TokenType.EndSub OrElse
                            tokenEnum.Current.Type = TokenType.Function OrElse
                            tokenEnum.Current.Type = TokenType.EndFunction Then
                    Exit While
                End If

                ifStatement.ThenStatements.Add(GetStatementFromTokens(tokenEnum))
                tokenEnum = ReadNextLine()
            End While

            If foundElseIf Then
                While tokenEnum IsNot Nothing
                    If tokenEnum.Current.Type = TokenType.EndIf Then
                        ifStatement.EndIfToken = tokenEnum.Current
                        foundEndIf = True
                        Exit While
                    End If

                    If tokenEnum.Current.Type = TokenType.Else Then
                        ifStatement.ElseToken = tokenEnum.Current
                        foundElse = True
                        Exit While
                    End If

                    If tokenEnum.Current.Type = TokenType.Sub OrElse tokenEnum.Current.Type = TokenType.EndSub OrElse
                              tokenEnum.Current.Type = TokenType.Function OrElse tokenEnum.Current.Type = TokenType.EndFunction OrElse
                              tokenEnum.Current.Type <> TokenType.ElseIf Then
                        Exit While
                    End If

                    ifStatement.ElseIfStatements.Add(ConstructElseIfStatement(tokenEnum))
                End While
            End If

            If foundElse Then
                tokenEnum.MoveNext()
                ExpectEndOfLine(tokenEnum)
                tokenEnum = ReadNextLine()

                While tokenEnum IsNot Nothing

                    If tokenEnum.Current.Type = TokenType.EndIf Then
                        ifStatement.EndIfToken = tokenEnum.Current
                        foundEndIf = True
                        Exit While
                    End If

                    If tokenEnum.Current.Type = TokenType.Sub OrElse
                                tokenEnum.Current.Type = TokenType.EndSub OrElse
                                tokenEnum.Current.Type = TokenType.Function OrElse
                                tokenEnum.Current.Type = TokenType.EndFunction OrElse
                                tokenEnum.Current.Type = TokenType.Else Then
                        Exit While
                    End If

                    ifStatement.ElseStatements.Add(GetStatementFromTokens(tokenEnum))
                    tokenEnum = ReadNextLine()
                End While
            End If

            If foundEndIf Then
                tokenEnum.MoveNext()
                ExpectEndOfLine(tokenEnum)
            Else
                AddError(ResourceHelper.GetString("EndIfExpected"))
            End If

            Return ifStatement
        End Function

        Private Function ConstructElseIfStatement(ByRef tokenEnum As TokenEnumerator) As ElseIfStatement
            Dim elseIfStatement As New ElseIfStatement() With {
                    .StartToken = tokenEnum.Current,
                    .ElseIfToken = tokenEnum.Current
            }

            tokenEnum.MoveNext()

            If EatLogicalExpression(tokenEnum, elseIfStatement.Condition) AndAlso
                        EatToken(tokenEnum, TokenType.Then, elseIfStatement.ThenToken) Then
                ExpectEndOfLine(tokenEnum)
            End If

            tokenEnum = ReadNextLine()

            While tokenEnum IsNot Nothing AndAlso
                        tokenEnum.Current.Type <> TokenType.EndIf AndAlso
                        tokenEnum.Current.Type <> TokenType.Else AndAlso
                        tokenEnum.Current.Type <> TokenType.ElseIf AndAlso
                        tokenEnum.Current.Type <> TokenType.Sub AndAlso
                        tokenEnum.Current.Type <> TokenType.EndSub AndAlso
                        tokenEnum.Current.Type <> TokenType.Function AndAlso
                        tokenEnum.Current.Type <> TokenType.EndFunction

                elseIfStatement.ThenStatements.Add(GetStatementFromTokens(tokenEnum))
                tokenEnum = ReadNextLine()
            End While

            Return elseIfStatement
        End Function

        Private Function ConstructSubStatement(tokenEnum As TokenEnumerator) As SubroutineStatement
            Dim subroutine As New SubroutineStatement() With {
                .StartToken = tokenEnum.Current,
                .SubToken = tokenEnum.Current
            }

            Dim tempRoutine = SubroutineStatement.Current
            SubroutineStatement.Current = subroutine

            tokenEnum.MoveNext()
            EatSimpleIdentifier(tokenEnum, subroutine.Name)

            If tokenEnum.Current.Type = TokenType.LeftParens Then
                tokenEnum.MoveNext()
                subroutine.Params = ParseParamList(tokenEnum, TokenType.RightParens)
                tokenEnum.MoveNext()
            End If

            ExpectEndOfLine(tokenEnum)
            tokenEnum = ReadNextLine()
            Dim flag = False

            While tokenEnum IsNot Nothing
                If tokenEnum.Current.Type = TokenType.EndSub OrElse tokenEnum.Current.Type = TokenType.EndFunction Then
                    subroutine.EndSubToken = tokenEnum.Current
                    flag = True
                    Exit While
                End If

                If tokenEnum.Current.Type = TokenType.Sub OrElse tokenEnum.Current.Type = TokenType.Function Then
                    RewindLine()
                    Exit While
                End If

                Dim statement = GetStatementFromTokens(tokenEnum)
                subroutine.Body.Add(statement)
                tokenEnum = ReadNextLine()
            End While

            SubroutineStatement.Current = tempRoutine

            If flag Then
                tokenEnum.MoveNext()
                ExpectEndOfLine(tokenEnum)
            Else
                AddError(ResourceHelper.GetString("EndSubExpected"))
            End If

            Return subroutine
        End Function

        Private Function ConstructGotoStatement(tokenEnum As TokenEnumerator) As GotoStatement
            Dim gotoStatement As New GotoStatement() With {
                .StartToken = tokenEnum.Current,
                .GotoToken = tokenEnum.Current,
                .subroutine = SubroutineStatement.Current
            }
            tokenEnum.MoveNext()
            EatSimpleIdentifier(tokenEnum, gotoStatement.Label)
            ExpectEndOfLine(tokenEnum)
            Return gotoStatement
        End Function

        Private Function ConstructLabelStatement(tokenEnum As TokenEnumerator) As LabelStatement
            Dim labelStatement As New LabelStatement() With {
                .StartToken = tokenEnum.Current,
                .LabelToken = tokenEnum.Current,
                .subroutine = SubroutineStatement.Current
            }
            tokenEnum.MoveNext()

            If EatToken(tokenEnum, TokenType.Colon, labelStatement.ColonToken) Then
                ExpectEndOfLine(tokenEnum)
            End If

            Return labelStatement
        End Function

        Private Function ConstructIdentifierStatement(tokenEnum As TokenEnumerator) As Statement
            Dim current = tokenEnum.Current
            Dim token As Token = tokenEnum.PeekNext()

            If token.Type = TokenType.Illegal Then
                AddError(tokenEnum.Current, ResourceHelper.GetString("UnrecognizedStatement"))
                Dim leftValue = BuildIdentifierTerm(tokenEnum)
                Return New AssignmentStatement() With {
                    .StartToken = current,
                    .LeftValue = leftValue
                }
            End If

            If token.Type = TokenType.Colon Then
                Return ConstructLabelStatement(tokenEnum)
            End If

            If token.Type = TokenType.LeftParens Then
                Return ConstructSubroutineCallStatement(tokenEnum)
            End If

            Dim expression = BuildIdentifierTerm(tokenEnum)

            If TypeOf expression Is MethodCallExpression Then
                ExpectEndOfLine(tokenEnum)
                Dim m = CType(expression, MethodCallExpression)
                Return New MethodCallStatement() With {
                    .StartToken = current,
                    .MethodCallExpression = m
                }
            End If

            Dim assignStatement As New AssignmentStatement() With {.StartToken = current, .LeftValue = expression}

            If EatOptionalToken(tokenEnum, TokenType.Equals, assignStatement.EqualsToken) Then
                assignStatement.RightValue = BuildArithmeticExpression(tokenEnum)

                If assignStatement.RightValue Is Nothing Then
                    If commaLine = -2 Then Return Nothing
                    AddError(tokenEnum.Current, ResourceHelper.GetString("ExpressionExpected"))
                End If

                ExpectEndOfLine(tokenEnum)

            Else
                AddError(tokenEnum.Current, ResourceHelper.GetString("UnrecognizedStatement"))
            End If

            Return assignStatement
        End Function

        Private Function ConstructSubroutineCallStatement(tokenEnum As TokenEnumerator) As Statement
            Dim subroutineCall As New SubroutineCallStatement() With {
                .StartToken = tokenEnum.Current,
                .Name = tokenEnum.Current,
                .OuterSubroutine = SubroutineStatement.Current
            }

            tokenEnum.MoveNext()

            Dim openP = tokenEnum.Current
            If EatToken(tokenEnum, TokenType.LeftParens) Then
                Dim caller = (subroutineCall.Name.Line, subroutineCall.Name.EndColumn)
                If openP.Line = commaLine AndAlso openP.Column = commaColumn Then
                    callerInfo = New CallerInfo(caller.Line, caller.EndColumn, 0)
                    commaLine = -2
                    Return Nothing
                End If

                Dim closeTokenFound = False
                subroutineCall.Args = ParseCommaSepaeatedList(tokenEnum, TokenType.RightParens, closeTokenFound, caller, False)
                If commaLine = -2 Then Return Nothing
                If closeTokenFound Then
                    tokenEnum.MoveNext()
                    ExpectEndOfLine(tokenEnum)
                End If
            End If

            Return subroutineCall
        End Function

        Public Function BuildArithmeticExpression(tokenEnum As TokenEnumerator) As Expression
            Return BuildExpression(tokenEnum, includeLogical:=True)
        End Function

        Public Function BuildLogicalExpression(tokenEnum As TokenEnumerator) As Expression
            Return BuildExpression(tokenEnum, includeLogical:=True)
        End Function

        Private Function BuildExpression(tokenEnum As TokenEnumerator, includeLogical As Boolean) As Expression
            If tokenEnum.IsEnd Then
                Return Nothing
            End If

            Dim current = tokenEnum.Current

            Dim leftHandExpr = BuildTerm(tokenEnum, includeLogical)
            If leftHandExpr Is Nothing Then Return Nothing

            While IsValidOperator(tokenEnum.Current.Type, includeLogical)
                Dim current2 = tokenEnum.Current
                tokenEnum.MoveNext()
                Dim rightHandExpr = BuildTerm(tokenEnum, includeLogical)
                If rightHandExpr Is Nothing Then Return Nothing

                leftHandExpr = MergeExpression(leftHandExpr, rightHandExpr, current2)
                If leftHandExpr Is Nothing Then Return Nothing
            End While

            leftHandExpr.StartToken = current
            leftHandExpr.EndToken = tokenEnum.Current
            Return leftHandExpr
        End Function

        Private Function BuildTerm(tokenEnum As TokenEnumerator, includeLogical As Boolean) As Expression
            Dim current = tokenEnum.Current

            If tokenEnum.IsEnd OrElse tokenEnum.Current.Type = TokenType.Illegal Then
                Return Nothing
            End If

            Dim expression As Expression
            Select Case tokenEnum.Current.Type
                Case TokenType.StringLiteral, TokenType.NumericLiteral, TokenType.DateLiteral, TokenType.True, TokenType.False
                    expression = New LiteralExpression(tokenEnum.Current)
                    expression.Precedence = 9
                    tokenEnum.MoveNext()

                Case TokenType.LeftParens
                    tokenEnum.MoveNext()
                    expression = BuildExpression(tokenEnum, includeLogical)
                    If expression Is Nothing Then Return Nothing
                    expression.Precedence = 9
                    If Not EatToken(tokenEnum, TokenType.RightParens) Then Return Nothing

                Case TokenType.LeftCurlyBracket
                    tokenEnum.MoveNext()
                    Dim closeTokenFound = False
                    Dim initExpr As New InitializerExpression(
                        precedence:=9,
                        arguments:=ParseCommaSepaeatedList(tokenEnum, TokenType.RightCurlyBracket, closeTokenFound, Nothing, False)
                    )

                    If commaLine = -2 Then Return Nothing
                    expression = initExpr
                    If closeTokenFound Then tokenEnum.MoveNext()

                Case TokenType.Identifier
                    expression = BuildIdentifierTerm(tokenEnum)
                    If expression Is Nothing Then Return Nothing

                Case Else
                    If tokenEnum.Current.Type <> TokenType.Subtraction Then
                        Return Nothing
                    End If

                    tokenEnum.MoveNext()
                    Dim expression2 = BuildTerm(tokenEnum, includeLogical)
                    If expression2 Is Nothing Then Return Nothing

                    expression = New NegativeExpression() With {
                        .Negation = tokenEnum.Current,
                        .Expression = expression2,
                        .Precedence = 9
                    }
            End Select

            expression.StartToken = current
            expression.EndToken = tokenEnum.Current
            Return expression
        End Function

        Private Function MergeExpression(leftHandExpr As Expression, rightHandExpr As Expression, operatorToken As Token) As Expression
            If TypeOf leftHandExpr Is InitializerExpression OrElse TypeOf rightHandExpr Is InitializerExpression Then
                AddError(operatorToken, "Array initializer can't be used as an operand in binary operations")
                Return Nothing
            End If

            Dim operatorPriority = GetOperatorPriority(operatorToken.Type)

            If operatorPriority <= leftHandExpr.Precedence Then
                Return New BinaryExpression() With {
                    .Operator = operatorToken,
                    .Precedence = operatorPriority,
                    .StartToken = leftHandExpr.StartToken,
                    .EndToken = rightHandExpr.EndToken,
                    .LeftHandSide = leftHandExpr,
                    .RightHandSide = rightHandExpr
                }
            End If

            Dim binaryExpression As BinaryExpression = TryCast(leftHandExpr, BinaryExpression)

            If binaryExpression IsNot Nothing Then
                Dim rightHandSide = binaryExpression.RightHandSide
                Dim expression As Expression

                If TypeOf rightHandSide Is BinaryExpression Then
                    expression = MergeExpression(rightHandSide, rightHandExpr, operatorToken)
                Else
                    expression = New BinaryExpression() With {
                        .Operator = operatorToken,
                        .Precedence = operatorPriority,
                        .StartToken = rightHandSide.StartToken,
                        .EndToken = rightHandExpr.EndToken,
                        .LeftHandSide = rightHandSide,
                        .RightHandSide = rightHandExpr
                    }
                End If

                binaryExpression.RightHandSide = expression
                binaryExpression.EndToken = rightHandExpr.EndToken
                Return binaryExpression
            End If

            Return Nothing
        End Function

        Private Function BuildIdentifierTerm(tokenEnum As TokenEnumerator) As Expression
            Dim current = tokenEnum.Current
            tokenEnum.MoveNext()
            Dim variable As Token = Nothing

            Dim tokenType = tokenEnum.Current.Type
            Dim openP As Token

            If tokenType = TokenType.Dot OrElse tokenType = TokenType.Lookup Then
                tokenEnum.MoveNext()

                If Not EatSimpleIdentifier(tokenEnum, variable) Then
                    Return Nothing
                End If

                openP = tokenEnum.Current
                If EatOptionalToken(tokenEnum, TokenType.LeftParens) Then
                    If tokenType = TokenType.Lookup Then
                        AddError(tokenEnum.Current, "The ! operator can't be used to call methods")
                    End If

                    Dim caller = (variable.Line, variable.EndColumn)
                    If openP.Line = commaLine AndAlso openP.Column = commaColumn Then
                        callerInfo = New CallerInfo(caller.Line, caller.EndColumn, 0)
                        commaLine = -2
                        Return Nothing
                    End If

                    Dim closeTokenFound = False
                    Dim methodCallExpression As New MethodCallExpression(
                        startToken:=current,
                        precedence:=9,
                        typeName:=current,
                        methodName:=variable,
                        arguments:=ParseCommaSepaeatedList(tokenEnum, TokenType.RightParens, closeTokenFound, caller, False)
                   )

                    ValidateFormName(methodCallExpression)

                    If commaLine = -2 Then Return Nothing

                    methodCallExpression.OuterSubroutine = SubroutineStatement.Current
                    methodCallExpression.EndToken = tokenEnum.Current
                    If closeTokenFound Then tokenEnum.MoveNext()
                    Return methodCallExpression
                End If

                Return New PropertyExpression() With {
                    .StartToken = current,
                    .EndToken = variable,
                    .Precedence = 9,
                    .TypeName = current,
                    .PropertyName = variable,
                    .IsDynamic = (tokenType = TokenType.Lookup)
                }
            End If

            If tokenEnum.Current.Type = TokenType.LeftBracket Then
                Dim leftHand As Expression = New IdentifierExpression() With {
                    .StartToken = current,
                    .Identifier = current,
                    .EndToken = current,
                    .Precedence = 9,
                    .Subroutine = SubroutineStatement.Current
                }

                Dim indexerExpression As Expression = Nothing

                While True
                    EatToken(tokenEnum, TokenType.LeftBracket)
                    indexerExpression = BuildArithmeticExpression(tokenEnum)

                    If indexerExpression Is Nothing Then
                        If commaLine = -2 Then Return Nothing
                        AddError(tokenEnum.Current, ResourceHelper.GetString("ExpressionExpected"))
                        Exit While
                    End If

                    EatToken(tokenEnum, TokenType.RightBracket)

                    If tokenEnum.Current.Type <> TokenType.LeftBracket Then
                        Exit While
                    End If

                    Dim arrayExpression As New ArrayExpression() With {
                        .StartToken = current,
                        .EndToken = tokenEnum.Current,
                        .Precedence = 9,
                        .LeftHand = leftHand,
                        .Indexer = indexerExpression
                    }
                    leftHand = arrayExpression
                End While

                Return New ArrayExpression() With {
                    .StartToken = current,
                    .EndToken = tokenEnum.Current,
                    .Precedence = 9,
                    .LeftHand = leftHand,
                    .Indexer = indexerExpression
                }
            End If

            openP = tokenEnum.Current
            If EatOptionalToken(tokenEnum, TokenType.LeftParens) Then
                Dim caller = (current.Line, current.EndColumn)
                If openP.Line = commaLine AndAlso openP.Column = commaColumn Then
                    callerInfo = New CallerInfo(caller.Line, caller.EndColumn, 0)
                    commaLine = -2
                    Return Nothing
                End If

                Dim closeTokenFound = False
                Dim methodCallExpression As New MethodCallExpression(
                        startToken:=current,
                        precedence:=9,
                        typeName:=Token.Illegal,
                        methodName:=current,
                        arguments:=ParseCommaSepaeatedList(tokenEnum, TokenType.RightParens, closeTokenFound, caller, False)
                )

                ValidateFormName(methodCallExpression)

                If commaLine = -2 Then Return Nothing

                methodCallExpression.OuterSubroutine = SubroutineStatement.Current
                methodCallExpression.EndToken = tokenEnum.Current
                If closeTokenFound Then tokenEnum.MoveNext()
                _FunctionsCall.Add(methodCallExpression)
                Return methodCallExpression
            End If

            Return New IdentifierExpression() With {
                .StartToken = current,
                .EndToken = current,
                .Precedence = 9,
                .Identifier = current,
                .Subroutine = SubroutineStatement.Current
            }
        End Function

        Private Sub ValidateFormName(m As MethodCallExpression)
            If FormNames Is Nothing Then Return

            If m.TypeName.LCaseText = "forms" Then
                Select Case m.MethodName.LCaseText
                    Case "showform", "showdialog"
                        If m.Arguments.Count > 0 Then
                            Dim strLit = TryCast(m.Arguments(0), LiteralExpression)
                            If strLit?.Literal.Type = TokenType.StringLiteral Then
                                Dim formName = strLit.Literal
                                If Not FormNames.Contains(formName.LCaseText.Trim("""")) Then
                                    AddError(strLit.Literal, $"The project doesn't contain a form named {formName.Text}")
                                End If
                            End If
                        End If
                End Select
            End If
        End Sub

        Public Shared Function ParseArgumentList(
                          args As String,
                          lineNumber As Integer,
                          lines As List(Of String),
                          openToken As TokenType
                    ) As List(Of Expression)

            Dim closeToken As TokenType
            Select Case openToken
                Case TokenType.LeftParens
                    args = "(" + args
                    closeToken = TokenType.RightParens
                Case TokenType.LeftBracket
                    args = "[" + args
                    closeToken = TokenType.RightBracket
                Case TokenType.LeftCurlyBracket
                    args = "{" + args
                    closeToken = TokenType.RightCurlyBracket
            End Select

            Dim tokens = LineScanner.GetTokenEnumerator(args, lineNumber, lines)
            tokens.MoveNext()
            Dim parser As New Parser()
            Dim closeTokenFound = False
            Dim argExprs = parser.ParseCommaSepaeatedList(tokens, closeToken, closeTokenFound, Nothing, False)
            If closeTokenFound Then tokens.MoveNext()
            If parser.Errors.Count > 0 Then Return Nothing
            Return argExprs
        End Function

        Private Function ParseCommaSepaeatedList(
                           tokenEnum As TokenEnumerator,
                           closeToken As TokenType,
                           <Out> ByRef closeTokenFound As Boolean,
                           caller? As (Line As Integer, EndColumn As Integer),
                           Optional commaIsOptional As Boolean = True
                  ) As List(Of Expression)

            Dim items As New List(Of Expression)
            closeTokenFound = False

            Do
                Dim current = tokenEnum.Current
                If current.Type = closeToken Then
                    If current.Line = commaLine AndAlso current.Column = commaColumn Then
                        If caller.HasValue Then
                            callerInfo = New CallerInfo(caller.Value.Line, caller.Value.EndColumn, items.Count - 1)
                            commaLine = -2
                            Return Nothing
                        End If
                    End If
                    closeTokenFound = True
                    Exit Do
                End If

                Dim expression = BuildArithmeticExpression(tokenEnum)

                If expression Is Nothing Then
                    If commaLine = -2 Then Return Nothing
                    AddError(tokenEnum.Current, ResourceHelper.GetString("ExpressionExpected"))
                    Exit Do
                End If

                items.Add(expression)
                If tokenEnum.IsEndOrComment Then Exit Do

                Dim comma = tokenEnum.Current
                If EatOptionalToken(tokenEnum, TokenType.Comma) Then
                    If comma.Line = commaLine AndAlso comma.Column = commaColumn Then
                        If caller.HasValue Then
                            callerInfo = New CallerInfo(caller.Value.Line, caller.Value.EndColumn, items.Count)
                            commaLine = -2
                            Return Nothing
                        End If
                    End If

                ElseIf Not commaIsOptional Then
                    current = tokenEnum.Current
                    Select Case current.Type
                        Case closeToken
                            If current.Line = commaLine AndAlso current.Column = commaColumn Then
                                If caller.HasValue Then
                                    callerInfo = New CallerInfo(caller.Value.Line, caller.Value.EndColumn, items.Count - 1)
                                    commaLine = -2
                                    Return Nothing
                                End If
                            End If
                            closeTokenFound = True

                        Case TokenType.RightBracket, TokenType.RightCurlyBracket, TokenType.RightParens
                            ' outer list is closed, and the inner list misses the closeToken
                        Case Else
                            TokenIsExpected(tokenEnum, TokenType.Comma)
                    End Select

                    Exit Do
                End If

                If tokenEnum.IsEndOrComment Then
                    If closeTokenFound Then
                        AddError(tokenEnum.Current, ResourceHelper.GetString("UnexpectedMethodCallEOL"))
                    End If
                    Exit Do
                End If
            Loop

            If Not closeTokenFound Then
                TokenIsExpected(tokenEnum, closeToken)
            End If
            Return items
        End Function

        Private Sub TokenIsExpected(tokenEnum As TokenEnumerator, expectedToken As TokenType)
            Dim current = tokenEnum.Current
            Dim token As New Token With {
                .Line = current.Line,
                .Column = If(tokenEnum.IsEnd, current.EndColumn, current.Column)
            }
            Dim expectedChar = ""
            Select Case expectedToken
                Case TokenType.RightBracket
                    expectedChar = "]"
                Case TokenType.RightCurlyBracket
                    expectedChar = "}"
                Case TokenType.RightParens
                    expectedChar = ")"
                Case TokenType.Comma
                    expectedChar = ","
            End Select
            AddError(token, $"`{expectedChar}` is expected here but not found")
        End Sub

        Private Function ParseParamList(tokenEnum As TokenEnumerator, closingToken As TokenType) As List(Of Token)
            Dim items As New List(Of Token)

            Do Until tokenEnum.Current.Type = closingToken

                If tokenEnum.Current.Type <> TokenType.Identifier Then
                    AddError(tokenEnum.Current, String.Format(ResourceHelper.GetString("UnexpectedTokenAtLocation"), tokenEnum.Current.Text))
                    Exit Do
                End If

                items.Add(tokenEnum.Current)
                If Not tokenEnum.MoveNext() Then
                    TokenIsExpected(tokenEnum, TokenType.RightParens)
                    Exit Do
                End If

                If Not EatOptionalToken(tokenEnum, TokenType.Comma) Then
                    If tokenEnum.Current.Type <> closingToken Then
                        TokenIsExpected(tokenEnum, TokenType.Comma)
                    End If
                    Exit Do
                End If

                If tokenEnum.IsEnd Then
                    AddError(tokenEnum.Current, ResourceHelper.GetString("UnexpectedTokenAtLocation"))
                    Exit Do
                End If
            Loop

            Return items
        End Function

        Private Function IsValidOperator(token As TokenType, includeLogical As Boolean) As Boolean
            Dim operatorPriority = GetOperatorPriority(token)

            If includeLogical Then
                Return operatorPriority > 0
            End If

            Return operatorPriority >= 7
        End Function

        Friend Shared Function GetOperatorPriority(token As TokenType) As Integer
            Select Case token
                Case TokenType.Division, TokenType.Multiplication
                    Return 8
                Case TokenType.Addition, TokenType.Subtraction
                    Return 7
                Case TokenType.LessThan, TokenType.LessThanEqualTo, TokenType.GreaterThan, TokenType.GreaterThanEqualTo
                    Return 6
                Case TokenType.Equals, TokenType.NotEqualTo
                    Return 5
                Case TokenType.And
                    Return 4
                Case TokenType.Or
                    Return 3
                Case Else
                    Return 0
            End Select
        End Function

        Public Sub New()
            Me.New(Nothing)
        End Sub

        Public Sub New(errors As List(Of [Error]))
            _Errors = errors
            If _Errors Is Nothing Then _Errors = New List(Of [Error])()
            _SymbolTable = New SymbolTable(_Errors)
        End Sub


        Public Sub Parse(
                        reader As TextReader,
                        Optional autoCompletion As Boolean = False
                   )

            If reader Is Nothing Then
                Throw New ArgumentNullException("reader")
            End If

            codeLines = New List(Of String)
            Do
                Dim line = reader.ReadLine
                If line Is Nothing Then Exit Do
                codeLines.Add(line)
            Loop

            Parse(autoCompletion)
        End Sub

        Public Sub Parse(
                        codeLines As List(Of String),
                        Optional autoCompletion As Boolean = False,
                        Optional globalParser As Parser = Nothing
                   )

            Me.codeLines = codeLines
            Parse(autoCompletion, globalParser)
        End Sub

        Private Sub Parse(
                         autoCompletion As Boolean,
                         Optional globalParser As Parser = Nothing
                    )

            If keepSymbols Then
                keepSymbols = False
            Else
                _Errors.Clear()
                _SymbolTable.Reset()

                Dim dynamics1 = globalParser?.SymbolTable.Dynamics
                If dynamics1 IsNot Nothing Then
                    Dim dynamics2 = _SymbolTable.Dynamics
                    For Each item In dynamics1
                        dynamics2.Add(item.Key, item.Value)
                    Next
                End If
            End If

            _SymbolTable.AutoCompletion = autoCompletion

            _ParseTree = New List(Of Statement)()

            Dim parentSub = SubroutineStatement.Current

            BuildParseTree()

            If commaLine <> -1 Then Return

            For Each statement In _ParseTree
                statement.Parent = parentSub
                statement.AddSymbols(_SymbolTable)
            Next

            If Not SymbolTable.IsLoweredCode Then
                For Each statement In _ParseTree
                    statement.InferType(_SymbolTable)
                Next
            End If

            For Each e In _SymbolTable.PossibleEventHandlers
                If _SymbolTable.Subroutines.ContainsKey(e.Id.LCaseText) Then
                    Dim id = _SymbolTable.AllIdentifiers(e.index)
                    id.SymbolType = Completion.CompletionItemType.SubroutineName
                    _SymbolTable.AllIdentifiers(e.index) = id

                    Dim token = _SymbolTable.Subroutines(e.Id.LCaseText)
                    Dim subroutine = CType(token.Parent, SubroutineStatement)
                    If subroutine.SubToken.Type = TokenType.Function Then
                        AddError(e.Id, $"Functions can't be used as event handlers.")
                    End If

                Else
                    AddError(e.Id, $"Subroutine `{e.Id.Text}` is not defiend.")
                End If
            Next

            For Each funcCall In _FunctionsCall
                Dim funcName = funcCall.MethodName.LCaseText
                If _SymbolTable.Subroutines.ContainsKey(funcName) Then
                    Dim subInfo = _SymbolTable.Subroutines(funcName)
                    If subInfo.Type = TokenType.Sub Then
                        AddError(
                            funcCall.MethodName,
                            String.Format(
                                    CultureInfo.CurrentUICulture,
                                    ResourceHelper.GetString("SubroutineEventAssignment"),
                                    funcCall.MethodName.Text
                           ) & vbCrLf & "       Don't add `( )` after the sub name when used as an event handler"
                       )
                    End If
                End If
            Next

            For Each identifier In _SymbolTable.AllIdentifiers
                If _SymbolTable.UsedBeforeDefind(identifier) Then
                    Errors.Add(New [Error](identifier, $"The variable `{identifier.Text}` is used before being initialized."))
                End If
            Next

            _SymbolTable.AutoCompletion = False
        End Sub

        Public Shared Function ParseDateLiteral(literal As String) As (Ticks As Long?, IsDate As Boolean)
            If literal.Length < 2 Then
                Return (Nothing, False)
            End If

            If literal(1) = "-" OrElse literal(1) = "+" Then
                Dim s As TimeSpan
                If TimeSpan.TryParse(literal.Trim("#"c, "+"c), CultureInfo.InvariantCulture, s) Then
                    Return (s.Ticks, False)
                Else
                    Return (Nothing, False)
                End If
            Else
                Dim d As Date
                If Date.TryParse(literal.Trim("#"), CultureInfo.InvariantCulture, DateTimeStyles.None, d) Then
                    Return (d.Ticks, True)
                Else
                    Return (Nothing, True)
                End If
            End If

        End Function


        Public Shared Function Parse(
                         code As String,
                         symbolTable As SymbolTable,
                         lineOffset As Integer
                   ) As Parser

            Dim _parser As New Parser() With {
                   ._SymbolTable = symbolTable,
                   .keepSymbols = True,
                   .lineOffset = lineOffset
            }

            symbolTable.IsLoweredCode = True
            _parser.Parse(New StringReader(code))
            symbolTable.IsLoweredCode = False

            Return _parser
        End Function

        Private Sub BuildParseTree()
            _FunctionsCall.Clear()
            _currentLine = -1

            While True
                Dim tokenEnum = ReadNextLine()
                If tokenEnum Is Nothing Then Exit While

                Dim statement = GetStatementFromTokens(tokenEnum)
                _ParseTree.Add(statement)
            End While
        End Sub

        Dim loweredLines() As String
        Dim loweredIndex As Integer = -1

        Friend Function ReadNextLine() As TokenEnumerator
            If _rewindRequested Then
                _rewindRequested = False
                Return _currentLineEnum
            End If

            LineScanner.SubLineComments.Clear()
            Dim line As String

            If loweredIndex > -1 AndAlso loweredIndex < loweredLines.Count - 1 Then
                loweredIndex += 1
                line = loweredLines(loweredIndex)

            Else
                loweredIndex = -1
                _currentLine += 1
                If _currentLine >= codeLines.Count Then Return Nothing
                line = codeLines(_currentLine)

                If line.Contains(vbLf) Then
                    ' This lines contains more than one line.
                    ' This happens due to lowering code.
                    ' These lines are combined not to change the source code lines numbers.
                    loweredLines = line.Split({vbLf}, StringSplitOptions.RemoveEmptyEntries)
                    loweredIndex = 0
                    line = loweredLines(0)
                End If
            End If

            _currentLineEnum = LineScanner.GetTokenEnumerator(
                line,
                _currentLine,
                codeLines,
                lineOffset
            )

            Return _currentLineEnum
        End Function

        Friend Sub RewindLine()
            _rewindRequested = True
        End Sub

        Friend Function GetStatementFromTokens(tokenEnum As TokenEnumerator) As Statement
            If tokenEnum.IsEnd Then
                Return New EmptyStatement() With {
                    .StartToken = New Token With {
                        .Line = tokenEnum.LineNumber
                    }
                }
            End If

            Dim statement As Statement = Nothing

            Select Case tokenEnum.Current.Type
                Case TokenType.While
                    statement = ConstructWhileStatement(tokenEnum)

                Case TokenType.For
                    statement = ConstructForStatement(tokenEnum)

                Case TokenType.ForEach
                    statement = ConstructForEachStatement(tokenEnum)

                Case TokenType.Goto
                    statement = ConstructGotoStatement(tokenEnum)

                Case TokenType.If
                    statement = ConstructIfStatement(tokenEnum)

                Case TokenType.ElseIf
                    AddError(tokenEnum.Current, String.Format(ResourceHelper.GetString("ElseIfUnexpected"), tokenEnum.Current.Text))
                    Return New IllegalStatement(tokenEnum.Current)

                Case TokenType.Sub, TokenType.Function
                    statement = ConstructSubStatement(tokenEnum)

                Case TokenType.Identifier
                    statement = ConstructIdentifierStatement(tokenEnum)
                    If commaLine = -2 Then Return Nothing

                Case TokenType.Comment
                    Dim emptyStatement As New EmptyStatement()
                    Dim comment = tokenEnum.Current
                    emptyStatement.StartToken = comment
                    statement = emptyStatement
                    _SymbolTable.AllCommentLines.Add(comment)

                Case TokenType.Return
                    Dim returnExpr As Expression = Nothing
                    Dim returnToken = tokenEnum.Current
                    tokenEnum.MoveNext()
                    If Not tokenEnum.IsEnd Then EatExpression(tokenEnum, returnExpr)
                    Dim subroutine = SubroutineStatement.Current
                    statement = New ReturnStatement With {
                        .StartToken = returnToken,
                        .ReturnExpression = returnExpr,
                        .Subroutine = subroutine
                    }

                    If subroutine Is Nothing Then
                        AddError(returnToken, "Return can only appear inside Sub and Function blocks")
                    ElseIf subroutine.SubToken.Type = TokenType.Sub AndAlso returnExpr IsNot Nothing Then
                        AddError(returnToken, "Sub routines can't return values")
                    End If

                Case TokenType.ExitLoop, TokenType.ContinueLoop
                    statement = New JumpLoopStatement With {
                        .StartToken = tokenEnum.Current,
                        .UpLevel = GetLevel(tokenEnum)
                    }
            End Select

            If statement Is Nothing Then
                AddError(tokenEnum.Current, String.Format(ResourceHelper.GetString("UnexpectedTokenAtLocation"), tokenEnum.Current.Text))
                Return New IllegalStatement(tokenEnum.Current)
            End If

            If tokenEnum.Current.Type = TokenType.Comment Then
                statement.EndingComment = tokenEnum.Current
            End If

            Return statement
        End Function

        Private Function GetLevel(tokenEnum As TokenEnumerator) As Integer
            tokenEnum.MoveNext()
            Dim upLevel = 0
            Do Until tokenEnum.IsEnd
                If tokenEnum.Current.Type = TokenType.Subtraction Then
                    If upLevel < 1000 Then
                        upLevel += 1
                    Else
                        AddError(tokenEnum.Current, String.Format(ResourceHelper.GetString("UnexpectedTokenAtLocation"), tokenEnum.Current.Text))
                        Exit Do
                    End If
                ElseIf tokenEnum.Current.Type = TokenType.Multiplication Then
                    If upLevel = 0 Then
                        upLevel = 1000
                    Else
                        AddError(tokenEnum.Current, String.Format(ResourceHelper.GetString("UnexpectedTokenAtLocation"), tokenEnum.Current.Text))
                        Exit Do
                    End If
                Else
                    AddError(tokenEnum.Current, String.Format(ResourceHelper.GetString("UnexpectedTokenAtLocation"), tokenEnum.Current.Text))
                    Exit Do
                End If
                tokenEnum.MoveNext()
            Loop
            Return upLevel
        End Function

        Friend Sub AddError(errorDescription As String)
            AddError(_currentLine, 0, errorDescription)
        End Sub

        Friend Sub AddError(line As Integer, subLine As Integer, errorDescription As String)
            AddError(line, subLine, 0, errorDescription)
        End Sub

        Friend Sub AddError(line As Integer, subLine As Integer, column As Integer, errorDescription As String)
            _Errors.Add(New [Error](line, subLine, column, errorDescription))
        End Sub

        Friend Sub AddError(token As Token, errorDescription As String)
            _Errors.Add(New [Error](token, errorDescription))
        End Sub

        Public Shared Function EvaluateExpression(expression As Expression) As Primitive
            Try
                Dim literalExpression = TryCast(expression, LiteralExpression)

                If literalExpression IsNot Nothing Then
                    Return New Primitive(literalExpression.Literal.Text)
                End If

                Dim negativeExpression As NegativeExpression = TryCast(expression, NegativeExpression)

                If negativeExpression IsNot Nothing Then
                    Return -EvaluateExpression(negativeExpression.Expression)
                End If

                Dim binaryExpression As BinaryExpression = TryCast(expression, BinaryExpression)

                If binaryExpression IsNot Nothing Then
                    Dim primitive = EvaluateExpression(binaryExpression.LeftHandSide)
                    Dim primitive2 = EvaluateExpression(binaryExpression.RightHandSide)

                    Select Case binaryExpression.Operator.Type
                        Case TokenType.Addition
                            Return primitive + primitive2
                        Case TokenType.Subtraction
                            Return primitive - primitive2
                        Case TokenType.Multiplication
                            Return primitive * primitive2
                        Case TokenType.Division
                            Return primitive / primitive2
                        Case TokenType.NotEqualTo
                            Return primitive.NotEqualTo(primitive2)
                        Case TokenType.Equals
                            Return primitive.EqualTo(primitive2)
                        Case TokenType.And
                            Return Primitive.op_And(primitive, primitive2)
                        Case TokenType.Or
                            Return Primitive.op_Or(primitive, primitive2)
                        Case TokenType.LessThan
                            Return primitive.LessThan(primitive2)
                        Case TokenType.LessThanEqualTo
                            Return primitive.LessThanOrEqualTo(primitive2)
                        Case TokenType.GreaterThan
                            Return primitive.GreaterThan(primitive2)
                        Case TokenType.GreaterThanEqualTo
                            Return primitive.GreaterThanOrEqualTo(primitive2)
                        Case TokenType.Dot, TokenType.Lookup, TokenType.LeftParens, TokenType.RightParens, TokenType.LeftBracket, TokenType.RightBracket
                    End Select
                End If

            Catch
            End Try

            Return 0
        End Function

        Friend Function EatToken(tokenEnum As TokenEnumerator, expectedToken As TokenType, <Out> ByRef token As Token) As Boolean
            If Not tokenEnum.IsEnd AndAlso tokenEnum.Current.Type = expectedToken Then
                token = tokenEnum.Current
                tokenEnum.MoveNext()
                Return True
            End If

            token = Token.Illegal
            AddError(tokenEnum.Current, String.Format(ResourceHelper.GetString("TokenExpected"), expectedToken))
            Return False
        End Function

        Friend Function EatToken(tokenEnum As TokenEnumerator, expectedToken As TokenType) As Boolean
            Dim token As Token
            Return EatToken(tokenEnum, expectedToken, token)
        End Function

        Friend Function EatOptionalToken(
                         tokenEnum As TokenEnumerator,
                         optionalToken As TokenType,
                         <Out> ByRef token As Token
                   ) As Boolean

            If Not tokenEnum.IsEnd AndAlso tokenEnum.Current.Type = optionalToken Then
                token = tokenEnum.Current
                tokenEnum.MoveNext()
                Return True
            End If

            token = Token.Illegal
            Return False
        End Function

        Friend Function EatOptionalToken(tokenEnum As TokenEnumerator, optionalToken As TokenType) As Boolean
            Dim token As Token
            Return EatOptionalToken(tokenEnum, optionalToken, token)
        End Function

        Friend Function EatSimpleIdentifier(
                          tokenEnum As TokenEnumerator,
                          <Out> ByRef token As Token
                    ) As Boolean

            If Not tokenEnum.IsEnd AndAlso tokenEnum.Current.Type = TokenType.Identifier Then
                token = tokenEnum.Current
                tokenEnum.MoveNext()
                If token.Text <> "_" Then Return True
            End If

            token = Token.Illegal
            AddError(tokenEnum.Current, ResourceHelper.GetString("IdentifierExpected"))
            Return False
        End Function


        Friend Function EatExpression(
                          tokenEnum As TokenEnumerator,
                          <Out> ByRef expression As Expression
                   ) As Boolean

            expression = BuildArithmeticExpression(tokenEnum)

            If expression IsNot Nothing Then Return True

            AddError(tokenEnum.Current, ResourceHelper.GetString("ExpressionExpected"))
            Return False
        End Function

        Friend Function EatLogicalExpression(tokenEnum As TokenEnumerator, <Out> ByRef expression As Expression) As Boolean
            expression = BuildLogicalExpression(tokenEnum)

            If expression IsNot Nothing Then
                Return True
            End If

            AddError(tokenEnum.Current, ResourceHelper.GetString("ConditionExpected"))
            Return False
        End Function

        Friend Function ExpectEndOfLine(tokenEnum As TokenEnumerator) As Boolean
            If tokenEnum.IsEndOrComment Then
                Return True
            Else
                AddError(tokenEnum.Current, String.Format(ResourceHelper.GetString("UnexpectedTokenFound"), tokenEnum.Current.Text))
                Return False
            End If
        End Function

    End Class

    Public Class CallerInfo
        Public Line As Integer
        Public EndColumn As Integer
        Public ParamIndex As Integer

        Public Sub New(line As Integer, endColumn As Integer, paramIndex As Integer)
            Me.Line = line
            Me.EndColumn = endColumn
            Me.ParamIndex = paramIndex
        End Sub
    End Class
End Namespace
