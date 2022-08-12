Imports System
Imports System.Collections.Generic
Imports System.IO
Imports Microsoft.SmallBasic.Expressions
Imports Microsoft.SmallBasic.Library
Imports Microsoft.SmallBasic.Statements
Imports System.Runtime.InteropServices
Imports System.Globalization

Namespace Microsoft.SmallBasic
    Public Class Parser
        Private codeLines As List(Of String)
        Private _currentLine As Integer
        Private _currentLineEnumerator As TokenEnumerator
        Private _rewindRequested As Boolean
        Private _FunctionsCall As New List(Of MethodCallExpression)

        Public ReadOnly Property TypeInfoBag As TypeInfoBag
        Public Property Errors As List(Of [Error])

        Public ReadOnly Property ParseTree As List(Of Statement)


        Public ReadOnly Property SymbolTable As SymbolTable

        Private Function ConstructWhileStatement(ByVal tokenEnumerator As TokenEnumerator) As WhileStatement
            Dim whileStatement As New WhileStatement()
            whileStatement.StartToken = tokenEnumerator.Current
            whileStatement.WhileToken = tokenEnumerator.Current
            Dim whileStatement2 = whileStatement
            tokenEnumerator.MoveNext()

            If EatLogicalExpression(tokenEnumerator, whileStatement2.Condition) Then
                ExpectEndOfLine(tokenEnumerator)
            End If

            tokenEnumerator = ReadNextLine()
            Dim flag = False

            While tokenEnumerator IsNot Nothing

                If tokenEnumerator.Current.Token = Token.EndWhile OrElse tokenEnumerator.Current.Token = Token.Wend Then
                    whileStatement2.EndLoopToken = tokenEnumerator.Current
                    flag = True
                    Exit While
                End If

                If tokenEnumerator.Current.Token = Token.Sub OrElse tokenEnumerator.Current.Token = Token.EndSub OrElse
                            tokenEnumerator.Current.Token = Token.Function OrElse tokenEnumerator.Current.Token = Token.EndFunction Then
                    Exit While
                End If

                whileStatement2.Body.Add(GetStatementFromTokens(tokenEnumerator))
                tokenEnumerator = ReadNextLine()
            End While

            If flag Then
                tokenEnumerator.MoveNext()
                ExpectEndOfLine(tokenEnumerator)
            Else
                AddError(ResourceHelper.GetString("EndWhileExpected"))
            End If

            Return whileStatement2
        End Function

        Private Function ConstructForStatement(ByVal tokenEnumerator As TokenEnumerator) As ForStatement
            Dim forStatement As New ForStatement() With {
                .StartToken = tokenEnumerator.Current,
                .ForToken = tokenEnumerator.Current,
                .Subroutine = SubroutineStatement.Current
            }

            tokenEnumerator.MoveNext()
            If EatSimpleIdentifier(tokenEnumerator, forStatement.Iterator) AndAlso
                    EatToken(tokenEnumerator, Token.Equals, forStatement.EqualsToken) AndAlso
                    EatExpression(tokenEnumerator, forStatement.InitialValue) AndAlso
                    EatToken(tokenEnumerator, Token.To, forStatement.ToToken) AndAlso
                    EatExpression(tokenEnumerator, forStatement.FinalValue) AndAlso
                    (Not EatOptionalToken(tokenEnumerator, Token.Step, forStatement.StepToken) OrElse
                        EatExpression(tokenEnumerator, forStatement.StepValue)) Then
                ExpectEndOfLine(tokenEnumerator)
            End If

            tokenEnumerator = ReadNextLine()
            Dim flag = False

            While tokenEnumerator IsNot Nothing

                If tokenEnumerator.Current.Token = Token.EndFor OrElse tokenEnumerator.Current.Token = Token.Next Then
                    forStatement.EndLoopToken = tokenEnumerator.Current
                    flag = True
                    Exit While
                End If

                If tokenEnumerator.Current.Token = Token.Sub OrElse tokenEnumerator.Current.Token = Token.EndSub OrElse
                        tokenEnumerator.Current.Token = Token.Function OrElse tokenEnumerator.Current.Token = Token.EndFunction Then
                    Exit While
                End If

                forStatement.Body.Add(GetStatementFromTokens(tokenEnumerator))
                tokenEnumerator = ReadNextLine()
            End While

            If flag Then
                tokenEnumerator.MoveNext()
                ExpectEndOfLine(tokenEnumerator)
            Else
                AddError(ResourceHelper.GetString("EndForExpected"))
            End If

            Return forStatement
        End Function

        Private Function ConstructIfStatement(tokenEnum As TokenEnumerator) As IfStatement
            Dim ifStatement As New IfStatement() With {
                    .StartToken = tokenEnum.Current,
                    .IfToken = tokenEnum.Current
            }

            tokenEnum.MoveNext()

            If EatLogicalExpression(tokenEnum, ifStatement.Condition) AndAlso
                        EatToken(tokenEnum, Token.Then, ifStatement.ThenToken) Then
                ExpectEndOfLine(tokenEnum)
            End If

            tokenEnum = ReadNextLine()
            Dim foundEndIf = False
            Dim foundElseIf = False
            Dim foundElse = False

            While tokenEnum IsNot Nothing

                If tokenEnum.Current.Token = Token.EndIf Then
                    ifStatement.EndIfToken = tokenEnum.Current
                    foundEndIf = True
                    Exit While
                End If

                If tokenEnum.Current.Token = Token.Else Then
                    ifStatement.ElseToken = tokenEnum.Current
                    foundElse = True
                    Exit While
                End If

                If tokenEnum.Current.Token = Token.ElseIf Then
                    foundElseIf = True
                    Exit While
                End If

                If tokenEnum.Current.Token = Token.Sub OrElse
                            tokenEnum.Current.Token = Token.EndSub OrElse
                            tokenEnum.Current.Token = Token.Function OrElse
                            tokenEnum.Current.Token = Token.EndFunction Then
                    Exit While
                End If

                ifStatement.ThenStatements.Add(GetStatementFromTokens(tokenEnum))
                tokenEnum = ReadNextLine()
            End While

            If foundElseIf Then
                While tokenEnum IsNot Nothing
                    If tokenEnum.Current.Token = Token.EndIf Then
                        ifStatement.EndIfToken = tokenEnum.Current
                        foundEndIf = True
                        Exit While
                    End If

                    If tokenEnum.Current.Token = Token.Else Then
                        ifStatement.ElseToken = tokenEnum.Current
                        foundElse = True
                        Exit While
                    End If

                    If tokenEnum.Current.Token = Token.Sub OrElse tokenEnum.Current.Token = Token.EndSub OrElse
                              tokenEnum.Current.Token = Token.Function OrElse tokenEnum.Current.Token = Token.EndFunction OrElse
                              tokenEnum.Current.Token <> Token.ElseIf Then
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

                    If tokenEnum.Current.Token = Token.EndIf Then
                        ifStatement.EndIfToken = tokenEnum.Current
                        foundEndIf = True
                        Exit While
                    End If

                    If tokenEnum.Current.Token = Token.Sub OrElse
                                tokenEnum.Current.Token = Token.EndSub OrElse
                                tokenEnum.Current.Token = Token.Function OrElse
                                tokenEnum.Current.Token = Token.EndFunction OrElse
                                tokenEnum.Current.Token = Token.Else Then
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
                        EatToken(tokenEnum, Token.Then, elseIfStatement.ThenToken) Then
                ExpectEndOfLine(tokenEnum)
            End If

            tokenEnum = ReadNextLine()

            While tokenEnum IsNot Nothing AndAlso
                        tokenEnum.Current.Token <> Token.EndIf AndAlso
                        tokenEnum.Current.Token <> Token.Else AndAlso
                        tokenEnum.Current.Token <> Token.ElseIf AndAlso
                        tokenEnum.Current.Token <> Token.Sub AndAlso
                        tokenEnum.Current.Token <> Token.EndSub AndAlso
                        tokenEnum.Current.Token <> Token.Function AndAlso
                        tokenEnum.Current.Token <> Token.EndFunction

                elseIfStatement.ThenStatements.Add(GetStatementFromTokens(tokenEnum))
                tokenEnum = ReadNextLine()
            End While

            Return elseIfStatement
        End Function

        Private Function ConstructSubStatement(tokenEnumerator As TokenEnumerator) As SubroutineStatement
            Dim subroutine As New SubroutineStatement() With {
                .StartToken = tokenEnumerator.Current,
                .SubToken = tokenEnumerator.Current
            }
            Dim tempRoutine = SubroutineStatement.Current
            SubroutineStatement.Current = subroutine

            tokenEnumerator.MoveNext()
            EatSimpleIdentifier(tokenEnumerator, subroutine.Name)

            Dim params As List(Of String) = Nothing
            If tokenEnumerator.Current.Token = Token.LeftParens Then
                tokenEnumerator.MoveNext()
                subroutine.Params = ParseParamList(tokenEnumerator, Token.RightParens)
                tokenEnumerator.MoveNext()

                params = (From param In subroutine.Params
                          Select param.Text).ToList

                If params.Count > 0 Then
                    Dim paramDef As New System.Text.StringBuilder()
                    For Each param In params
                        paramDef.AppendLine($"{param} = Stack.PopValue(""_sVB_Arguments"")")
                    Next

                    Dim _parser = Parser.Parse(paramDef.ToString(), New SymbolTable(New List(Of [Error]), _TypeInfoBag), _TypeInfoBag)
                    For Each statment As AssignmentStatement In _parser._ParseTree
                        statment.IsLocalVariable = True
                        subroutine.Body.Add(statment)
                    Next
                End If
            End If

            ExpectEndOfLine(tokenEnumerator)
            tokenEnumerator = ReadNextLine()
            Dim flag = False


            Dim returnLabelToken As New TokenInfo() With {
                    .Text = $"_EXIT_SUB_{subroutine.Name.NormalizedText}",
                    .Token = Token.Identifier,
                    .TokenType = TokenType.Identifier
                }

            Dim returnLabelStatement As New LabelStatement() With {
                .ColonToken = New TokenInfo With {
                        .Text = ":",
                        .Token = Token.Colon,
                        .TokenType = TokenType.Delimiter
                },
                .StartToken = returnLabelToken,
                .LabelToken = returnLabelToken,
                .subroutine = subroutine
            }

            While tokenEnumerator IsNot Nothing
                If tokenEnumerator.Current.Token = Token.EndSub OrElse tokenEnumerator.Current.Token = Token.EndFunction Then
                    subroutine.EndSubToken = tokenEnumerator.Current
                    flag = True
                    Exit While
                End If

                If tokenEnumerator.Current.Token = Token.Sub OrElse tokenEnumerator.Current.Token = Token.Function Then
                    RewindLine()
                    Exit While
                End If

                Dim statement = GetStatementFromTokens(tokenEnumerator)
                subroutine.Body.Add(statement)
                tokenEnumerator = ReadNextLine()
            End While

            subroutine.Body.Add(returnLabelStatement)
            SubroutineStatement.Current = tempRoutine

            If flag Then
                tokenEnumerator.MoveNext()
                ExpectEndOfLine(tokenEnumerator)
            Else
                AddError(ResourceHelper.GetString("EndSubExpected"))
            End If

            Return subroutine
        End Function

        Private Function ConstructGotoStatement(ByVal tokenEnumerator As TokenEnumerator) As GotoStatement
            Dim gotoStatement As New GotoStatement() With {
                .StartToken = tokenEnumerator.Current,
                .GotoToken = tokenEnumerator.Current,
                .subroutine = SubroutineStatement.Current
            }
            tokenEnumerator.MoveNext()
            EatSimpleIdentifier(tokenEnumerator, gotoStatement.Label)
            ExpectEndOfLine(tokenEnumerator)
            Return gotoStatement
        End Function

        Private Function ConstructLabelStatement(ByVal tokenEnumerator As TokenEnumerator) As LabelStatement
            Dim labelStatement As New LabelStatement() With {
                .StartToken = tokenEnumerator.Current,
                .LabelToken = tokenEnumerator.Current,
                .subroutine = SubroutineStatement.Current
            }
            tokenEnumerator.MoveNext()

            If EatToken(tokenEnumerator, Token.Colon, labelStatement.ColonToken) Then
                ExpectEndOfLine(tokenEnumerator)
            End If

            Return labelStatement
        End Function

        Private Function ConstructIdentifierStatement(ByVal tokenEnumerator As TokenEnumerator) As Statement
            Dim current = tokenEnumerator.Current
            Dim tokenInfo As TokenInfo = tokenEnumerator.PeekNext()

            If tokenInfo.Token = Token.Illegal Then
                AddError(tokenEnumerator.Current, ResourceHelper.GetString("UnrecognizedStatement"))
                Dim leftValue = BuildIdentifierTerm(tokenEnumerator)
                Return New AssignmentStatement() With {
                    .StartToken = current,
                    .LeftValue = leftValue
                }
            End If

            If tokenInfo.Token = Token.Colon Then
                Return ConstructLabelStatement(tokenEnumerator)
            End If

            If tokenInfo.Token = Token.LeftParens Then
                Return ConstructSubroutineCallStatement(tokenEnumerator)
            End If

            Dim expression = BuildIdentifierTerm(tokenEnumerator)

            If TypeOf expression Is MethodCallExpression Then
                ExpectEndOfLine(tokenEnumerator)
                Dim methodCallStatement As New MethodCallStatement()
                methodCallStatement.StartToken = current
                methodCallStatement.MethodCallExpression = TryCast(expression, MethodCallExpression)
                Return methodCallStatement
            End If

            Dim assignStatement As New AssignmentStatement() With {.StartToken = current, .LeftValue = expression}

            If EatOptionalToken(tokenEnumerator, Token.Equals, assignStatement.EqualsToken) Then
                assignStatement.RightValue = BuildArithmeticExpression(tokenEnumerator)

                If assignStatement.RightValue Is Nothing Then
                    AddError(tokenEnumerator.Current, ResourceHelper.GetString("ExpressionExpected"))
                End If

                ExpectEndOfLine(tokenEnumerator)

            Else
                AddError(tokenEnumerator.Current, ResourceHelper.GetString("UnrecognizedStatement"))
            End If

            Return assignStatement
        End Function

        Private Function ConstructSubroutineCallStatement(tokenEnumerator As TokenEnumerator) As Statement
            Dim subroutineCall As New SubroutineCallStatement() With {
                .StartToken = tokenEnumerator.Current,
                .Name = tokenEnumerator.Current,
                .OuterSubRoutine = SubroutineStatement.Current
            }

            tokenEnumerator.MoveNext()

            If EatToken(tokenEnumerator, Token.LeftParens) Then
                subroutineCall.Args = ParseCommaSepList(tokenEnumerator, Token.RightParens, False)
                If EatToken(tokenEnumerator, Token.RightParens) Then ExpectEndOfLine(tokenEnumerator)
            End If

            Return subroutineCall
        End Function

        Public Function BuildArithmeticExpression(tokenEnumerator As TokenEnumerator) As Expression
            Return BuildExpression(tokenEnumerator, includeLogical:=True)
        End Function

        Public Function BuildLogicalExpression(ByVal tokenEnumerator As TokenEnumerator) As Expression
            Return BuildExpression(tokenEnumerator, includeLogical:=True)
        End Function

        Private Function BuildExpression(ByVal tokenEnumerator As TokenEnumerator, ByVal includeLogical As Boolean) As Expression
            If tokenEnumerator.IsEndOfList Then
                Return Nothing
            End If

            Dim current = tokenEnumerator.Current

            Dim leftHandExpr = BuildTerm(tokenEnumerator, includeLogical)
            If leftHandExpr Is Nothing Then Return Nothing

            While IsValidOperator(tokenEnumerator.Current.Token, includeLogical)
                Dim current2 = tokenEnumerator.Current
                tokenEnumerator.MoveNext()
                Dim rightHandExpr = BuildTerm(tokenEnumerator, includeLogical)

                If rightHandExpr Is Nothing Then Return Nothing

                leftHandExpr = MergeExpression(leftHandExpr, rightHandExpr, current2)
                If leftHandExpr Is Nothing Then Return Nothing
            End While

            leftHandExpr.StartToken = current
            leftHandExpr.EndToken = tokenEnumerator.Current
            Return leftHandExpr
        End Function

        Private Function BuildTerm(tokenEnumerator As TokenEnumerator, includeLogical As Boolean) As Expression
            Dim current = tokenEnumerator.Current

            If tokenEnumerator.IsEndOfList OrElse tokenEnumerator.Current.Token = Token.Illegal Then
                Return Nothing
            End If

            Dim expression As Expression
            Select Case tokenEnumerator.Current.Token
                Case Token.StringLiteral, Token.NumericLiteral, Token.True, Token.False
                    expression = New LiteralExpression(tokenEnumerator.Current)
                    expression.Precedence = 9
                    tokenEnumerator.MoveNext()

                Case Token.LeftParens
                    tokenEnumerator.MoveNext()
                    expression = BuildExpression(tokenEnumerator, includeLogical)
                    If expression Is Nothing Then Return Nothing
                    expression.Precedence = 9
                    If Not EatToken(tokenEnumerator, Token.RightParens) Then Return Nothing

                Case Token.LeftCurlyBracket
                    tokenEnumerator.MoveNext()
                    Dim initExpr As New InitializerExpression(
                        precedence:=9,
                        arguments:=ParseCommaSepList(tokenEnumerator, Token.RightCurlyBracket, False)
                   )

                    If initExpr.Arguments.Count = 0 Then Return Nothing
                    expression = initExpr
                    tokenEnumerator.MoveNext() ' }

                Case Token.Identifier
                    expression = BuildIdentifierTerm(tokenEnumerator)
                    If expression Is Nothing Then Return Nothing

                Case Else
                    If tokenEnumerator.Current.Token <> Token.Subtraction Then
                        Return Nothing
                    End If

                    tokenEnumerator.MoveNext()
                    Dim expression2 = BuildTerm(tokenEnumerator, includeLogical)

                    If expression2 Is Nothing Then
                        Return Nothing
                    End If

                    expression = New NegativeExpression() With {
                        .Negation = tokenEnumerator.Current,
                        .Expression = expression2,
                        .Precedence = 9
                    }
            End Select

            expression.StartToken = current
            expression.EndToken = tokenEnumerator.Current
            Return expression
        End Function

        Private Function MergeExpression(ByVal leftHandExpr As Expression, ByVal rightHandExpr As Expression, ByVal operatorToken As TokenInfo) As Expression
            If TypeOf leftHandExpr Is InitializerExpression OrElse TypeOf rightHandExpr Is InitializerExpression Then
                AddError(operatorToken, "Array initializer can't be used as an operand in binary operations")
                Return Nothing
            End If

            Dim operatorPriority = GetOperatorPriority(operatorToken.Token)

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
                Dim expression As Expression = Nothing

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

        Private Function BuildIdentifierTerm(ByVal tokenEnumerator As TokenEnumerator) As Expression
            Dim current = tokenEnumerator.Current
            tokenEnumerator.MoveNext()
            Dim variable As TokenInfo = Nothing

            Dim t = tokenEnumerator.Current.Token
            If t = Token.Dot OrElse t = Token.Lookup Then
                tokenEnumerator.MoveNext()

                If Not EatSimpleIdentifier(tokenEnumerator, variable) Then
                    Return Nothing
                End If

                If EatOptionalToken(tokenEnumerator, Token.LeftParens) Then
                    If t = Token.Lookup Then
                        AddError(tokenEnumerator.Current, "The ! operator can't be used to call methods")
                    End If

                    Dim methodCallExpression As New MethodCallExpression(
                        startToken:=current,
                        precedence:=9,
                        typeName:=current,
                        methodName:=variable,
                        arguments:=ParseCommaSepList(tokenEnumerator, Token.RightParens, False)
                   )

                    methodCallExpression.OuterSubRoutine = SubroutineStatement.Current
                    methodCallExpression.EndToken = tokenEnumerator.Current
                    tokenEnumerator.MoveNext()
                    Return methodCallExpression
                End If

                Return New PropertyExpression() With {
                    .StartToken = current,
                    .EndToken = variable,
                    .Precedence = 9,
                    .TypeName = current,
                    .PropertyName = variable,
                    .IsDynamic = (t = Token.Lookup)
                }
            End If

            If tokenEnumerator.Current.Token = Token.LeftBracket Then
                Dim leftHand As Expression = New IdentifierExpression() With {
                    .StartToken = current,
                    .Identifier = current,
                    .EndToken = current,
                    .Precedence = 9,
                    .Subroutine = SubroutineStatement.Current
                }

                Dim indexerExpression As Expression = Nothing

                While True
                    EatToken(tokenEnumerator, Token.LeftBracket)
                    indexerExpression = BuildArithmeticExpression(tokenEnumerator)

                    If indexerExpression Is Nothing Then
                        AddError(tokenEnumerator.Current, ResourceHelper.GetString("ExpressionExpected"))
                        Exit While
                    End If

                    EatToken(tokenEnumerator, Token.RightBracket)

                    If tokenEnumerator.Current.Token <> Token.LeftBracket Then
                        Exit While
                    End If

                    Dim arrayExpression As New ArrayExpression() With {
                        .StartToken = current,
                        .EndToken = tokenEnumerator.Current,
                        .Precedence = 9,
                        .LeftHand = leftHand,
                        .Indexer = indexerExpression
                    }
                    leftHand = arrayExpression
                End While

                Return New ArrayExpression() With {
                    .StartToken = current,
                    .EndToken = tokenEnumerator.Current,
                    .Precedence = 9,
                    .LeftHand = leftHand,
                    .Indexer = indexerExpression
                }
            End If

            If EatOptionalToken(tokenEnumerator, Token.LeftParens) Then
                Dim methodCallExpression As New MethodCallExpression(
                        startToken:=current,
                        precedence:=9,
                        typeName:=TokenInfo.Illegal,
                        methodName:=current,
                        arguments:=ParseCommaSepList(tokenEnumerator, Token.RightParens, False)
                )

                methodCallExpression.OuterSubRoutine = SubroutineStatement.Current
                methodCallExpression.EndToken = tokenEnumerator.Current
                tokenEnumerator.MoveNext()
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

        Public Shared Function ParseArgumentList(
                          args As String,
                          lineNumber As Integer,
                          lines As List(Of String),
                          openToken As Token
                    ) As List(Of Expression)

            Dim closeToken As Token
            Select Case openToken
                Case Token.LeftParens
                    args = "(" + args
                    closeToken = Token.RightParens
                Case Token.LeftBracket
                    args = "[" + args
                    closeToken = Token.RightBracket
                Case Token.LeftCurlyBracket
                    args = "{" + args
                    closeToken = Token.RightCurlyBracket
            End Select

            Dim tokens = LineScanner.GetTokenEnumerator(args, lineNumber, lines)
            tokens.MoveNext()
            Dim parser As New Parser()
            Dim argExprs = parser.ParseCommaSepList(tokens, closeToken, False)

            If parser.Errors.Count > 0 Then Return Nothing
            Return argExprs
        End Function

        Private Function ParseCommaSepList(tokenEnumerator As TokenEnumerator, closingToken As Token, Optional commaIsOptional As Boolean = True) As List(Of Expression)
            Dim items As New List(Of Expression)

            While tokenEnumerator.Current.Token <> closingToken
                Dim expression = BuildArithmeticExpression(tokenEnumerator)

                If expression Is Nothing Then
                    AddError(tokenEnumerator.Current, ResourceHelper.GetString("ExpressionExpected"))
                    Exit While
                End If

                items.Add(expression)
                If Not EatOptionalToken(tokenEnumerator, Token.Comma) Then
                    If Not commaIsOptional Then
                        If tokenEnumerator.Current.Token <> closingToken Then
                            AddError(tokenEnumerator.Current, "Comma is expected here but not found.")
                        End If
                        Exit While
                    End If
                End If

                If tokenEnumerator.IsEndOfNonCommentList Then
                    AddError(tokenEnumerator.Current, ResourceHelper.GetString("UnexpectedMethodCallEOL"))
                    Exit While
                End If
            End While

            Return items
        End Function

        Private Function ParseParamList(tokenEnumerator As TokenEnumerator, closingToken As Token) As List(Of TokenInfo)
            Dim items As New List(Of TokenInfo)


            While tokenEnumerator.Current.Token <> closingToken
                If tokenEnumerator.Current.Token <> Token.Identifier Then
                    AddError(tokenEnumerator.Current, String.Format(ResourceHelper.GetString("UnexpectedTokenAtLocation"), tokenEnumerator.Current.Text))
                    Exit While
                End If

                items.Add(tokenEnumerator.Current)
                tokenEnumerator.MoveNext()

                If Not EatOptionalToken(tokenEnumerator, Token.Comma) Then
                    If tokenEnumerator.Current.TokenType = TokenType.Comment Then
                        Dim nextToken = tokenEnumerator.PeekNext
                        If tokenEnumerator.Current.Token <> closingToken Then
                            AddError(tokenEnumerator.Current, "Comma is expected here but not found.")
                        End If
                        Exit While
                    End If
                End If

                If tokenEnumerator.IsEndOfNonCommentList Then
                    AddError(tokenEnumerator.Current, ResourceHelper.GetString("UnexpectedTokenAtLocation"))
                    Exit While
                End If
            End While

            Return items
        End Function

        Private Function IsValidOperator(ByVal token As Token, ByVal includeLogical As Boolean) As Boolean
            Dim operatorPriority = GetOperatorPriority(token)

            If includeLogical Then
                Return operatorPriority > 0
            End If

            Return operatorPriority >= 7
        End Function

        Friend Shared Function GetOperatorPriority(ByVal token As Token) As Integer
            Select Case token
                Case Token.Division, Token.Multiplication
                    Return 8
                Case Token.Addition, Token.Subtraction
                    Return 7
                Case Token.LessThan, Token.LessThanEqualTo, Token.GreaterThan, Token.GreaterThanEqualTo
                    Return 6
                Case Token.Equals, Token.NotEqualTo
                    Return 5
                Case Token.And
                    Return 4
                Case Token.Or
                    Return 3
                Case Else
                    Return 0
            End Select
        End Function

        Public Sub New()
            Me.New(Nothing, Nothing)
        End Sub

        Public Sub New(errors As List(Of [Error]), typeInfoBag As TypeInfoBag)
            _Errors = errors
            _TypeInfoBag = typeInfoBag
            If _Errors Is Nothing Then _Errors = New List(Of [Error])()
            _SymbolTable = New SymbolTable(_Errors, typeInfoBag)
        End Sub


        Dim keepSymbols As Boolean

        Public Sub Parse(reader As TextReader, Optional autoCompletion As Boolean = False)
            If reader Is Nothing Then
                Throw New ArgumentNullException("reader")
            End If


            _TypeInfoBag = TypeInfoBag

            If keepSymbols Then
                keepSymbols = False
            Else
                _Errors.Clear()
                _SymbolTable.Reset()
            End If

            _SymbolTable.autoCompletion = autoCompletion

            _ParseTree = New List(Of Statement)()
            codeLines = New List(Of String)
            Do
                Dim line = reader.ReadLine
                If line Is Nothing Then Exit Do
                codeLines.Add(line)
            Loop

            BuildParseTree()

            For Each item In _ParseTree
                item.AddSymbols(_SymbolTable)
            Next

            For Each funcCall In _FunctionsCall
                Dim funcName = funcCall.MethodName.NormalizedText
                If _SymbolTable.Subroutines.ContainsKey(funcName) Then
                    Dim subInfo = _SymbolTable.Subroutines(funcName)
                    If subInfo.Token = Token.Sub Then
                        AddError(funcCall.MethodName, String.Format(CultureInfo.CurrentUICulture, ResourceHelper.GetString("SubroutineEventAssignment"), funcCall.MethodName.Text))
                    End If
                End If
            Next
            _SymbolTable.autoCompletion = False
        End Sub

        Public Shared Function Parse(code As String, symbolTable As SymbolTable, typeInfoBag As TypeInfoBag) As Parser
            Dim _parser As New Parser()
            _parser._SymbolTable = symbolTable
            _parser._TypeInfoBag = typeInfoBag
            _parser.keepSymbols = True
            _parser.Parse(New IO.StringReader(code))
            Return _parser
        End Function

        Private Sub BuildParseTree()
            _currentLine = -1

            While True
                Dim tokenEnumerator = ReadNextLine()
                If tokenEnumerator Is Nothing Then Exit While

                Dim statement = GetStatementFromTokens(tokenEnumerator)
                _ParseTree.Add(statement)
            End While
        End Sub

        Friend Function ReadNextLine() As TokenEnumerator
            If _rewindRequested Then
                _rewindRequested = False
                Return _currentLineEnumerator
            End If

            _currentLine += 1
            If _currentLine >= codeLines.Count Then Return Nothing

            _currentLineEnumerator = LineScanner.GetTokenEnumerator(codeLines(_currentLine), _currentLine, codeLines)
            Return _currentLineEnumerator
        End Function

        Friend Sub RewindLine()
            _rewindRequested = True
        End Sub

        Friend Function GetStatementFromTokens(tokenEnumerator As TokenEnumerator) As Statement
            If tokenEnumerator.IsEndOfList Then
                Return New EmptyStatement() With {
                    .StartToken = New TokenInfo With {
                        .Line = tokenEnumerator.LineNumber
                    }
                }
            End If

            Dim statement As Statement = Nothing

            Select Case tokenEnumerator.Current.Token
                Case Token.While
                    statement = ConstructWhileStatement(tokenEnumerator)

                Case Token.For
                    statement = ConstructForStatement(tokenEnumerator)

                Case Token.Goto
                    statement = ConstructGotoStatement(tokenEnumerator)

                Case Token.If
                    statement = ConstructIfStatement(tokenEnumerator)

                Case Token.ElseIf
                    AddError(tokenEnumerator.Current, String.Format(ResourceHelper.GetString("ElseIfUnexpected"), tokenEnumerator.Current.Text))
                    Return New IllegalStatement(tokenEnumerator.Current)

                Case Token.Sub, Token.Function
                    statement = ConstructSubStatement(tokenEnumerator)

                Case Token.Identifier
                    statement = ConstructIdentifierStatement(tokenEnumerator)

                Case Token.Comment
                    Dim emptyStatement As New EmptyStatement()
                    emptyStatement.StartToken = tokenEnumerator.Current
                    statement = emptyStatement

                Case Token.Return
                    Dim returnExpr As Expression = Nothing
                    Dim returnToken = tokenEnumerator.Current
                    tokenEnumerator.MoveNext()
                    If Not tokenEnumerator.IsEndOfNonCommentList Then EatExpression(tokenEnumerator, returnExpr)
                    Dim subroutine = SubroutineStatement.Current
                    statement = New ReturnStatement With {
                        .StartToken = returnToken,
                        .ReturnExpression = returnExpr,
                        .Subroutine = subroutine
                    }

                    If subroutine Is Nothing Then
                        AddError(returnToken, "Return can only appear insid Sub and Function blocks")
                    ElseIf subroutine.SubToken.Token = Token.Sub AndAlso returnExpr IsNot Nothing Then
                        AddError(returnToken, "Sub routines can't return values")
                    End If

                Case Token.ExitLoop, Token.ContinueLoop
                    statement = New JumpLoopStatement With {
                        .StartToken = tokenEnumerator.Current,
                        .UpLevel = GetLevel(tokenEnumerator)
                    }
            End Select

            If statement Is Nothing Then
                AddError(tokenEnumerator.Current, String.Format(ResourceHelper.GetString("UnexpectedTokenAtLocation"), tokenEnumerator.Current.Text))
                Return New IllegalStatement(tokenEnumerator.Current)
            End If

            If tokenEnumerator.Current.Token = Token.Comment Then
                statement.EndingComment = tokenEnumerator.Current
            End If

            Return statement
        End Function

        Private Function GetLevel(tokenEnumerator As TokenEnumerator) As Integer
            tokenEnumerator.MoveNext()
            Dim upLevel = 0
            Do Until tokenEnumerator.IsEndOfNonCommentList
                If tokenEnumerator.Current.Token = Token.Subtraction Then
                    If upLevel < 1000 Then
                        upLevel += 1
                    Else
                        AddError(tokenEnumerator.Current, String.Format(ResourceHelper.GetString("UnexpectedTokenAtLocation"), tokenEnumerator.Current.Text))
                        Exit Do
                    End If
                ElseIf tokenEnumerator.Current.Token = Token.Multiplication Then
                    If upLevel = 0 Then
                        upLevel = 1000
                    Else
                        AddError(tokenEnumerator.Current, String.Format(ResourceHelper.GetString("UnexpectedTokenAtLocation"), tokenEnumerator.Current.Text))
                        Exit Do
                    End If
                Else
                    AddError(tokenEnumerator.Current, String.Format(ResourceHelper.GetString("UnexpectedTokenAtLocation"), tokenEnumerator.Current.Text))
                    Exit Do
                End If
                tokenEnumerator.MoveNext()
            Loop
            Return upLevel
        End Function

        Friend Sub AddError(ByVal errorDescription As String)
            AddError(_currentLine, 0, errorDescription)
        End Sub

        Friend Sub AddError(line As Integer, subLine As Integer, errorDescription As String)
            AddError(line, subLine, 0, errorDescription)
        End Sub

        Friend Sub AddError(line As Integer, subLine As Integer, column As Integer, errorDescription As String)
            _Errors.Add(New [Error](line, subLine, column, errorDescription))
        End Sub

        Friend Sub AddError(ByVal tokenInfo As TokenInfo, ByVal errorDescription As String)
            _Errors.Add(New [Error](tokenInfo, errorDescription))
        End Sub

        Public Shared Function EvaluateExpression(ByVal expression As Expression) As Primitive
            Try
                Dim literalExpression As LiteralExpression = TryCast(expression, LiteralExpression)

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

                    Select Case binaryExpression.Operator.Token
                        Case Token.Addition
                            Return primitive + primitive2
                        Case Token.Subtraction
                            Return primitive - primitive2
                        Case Token.Multiplication
                            Return primitive * primitive2
                        Case Token.Division
                            Return primitive / primitive2
                        Case Token.NotEqualTo
                            Return primitive.NotEqualTo(primitive2)
                        Case Token.Equals
                            Return primitive.EqualTo(primitive2)
                        Case Token.And
                            Return Primitive.op_And(primitive, primitive2)
                        Case Token.Or
                            Return Primitive.op_Or(primitive, primitive2)
                        Case Token.LessThan
                            Return primitive.LessThan(primitive2)
                        Case Token.LessThanEqualTo
                            Return primitive.LessThanOrEqualTo(primitive2)
                        Case Token.GreaterThan
                            Return primitive.GreaterThan(primitive2)
                        Case Token.GreaterThanEqualTo
                            Return primitive.GreaterThanOrEqualTo(primitive2)
                        Case Token.Dot, Token.Lookup, Token.LeftParens, Token.RightParens, Token.LeftBracket, Token.RightBracket
                    End Select
                End If

            Catch
            End Try

            Return 0
        End Function

        Friend Function EatToken(ByVal tokenEnumerator As TokenEnumerator, ByVal expectedToken As Token, <Out> ByRef tokenInfo As TokenInfo) As Boolean
            If Not tokenEnumerator.IsEndOfList AndAlso tokenEnumerator.Current.Token = expectedToken Then
                tokenInfo = tokenEnumerator.Current
                tokenEnumerator.MoveNext()
                Return True
            End If

            tokenInfo = TokenInfo.Illegal
            AddError(tokenEnumerator.Current, String.Format(ResourceHelper.GetString("TokenExpected"), expectedToken))
            Return False
        End Function

        Friend Function EatToken(ByVal tokenEnumerator As TokenEnumerator, ByVal expectedToken As Token) As Boolean
            Dim tokenInfo As TokenInfo
            Return EatToken(tokenEnumerator, expectedToken, tokenInfo)
        End Function

        Friend Function EatOptionalToken(ByVal tokenEnumerator As TokenEnumerator, ByVal optionalToken As Token, <Out> ByRef tokenInfo As TokenInfo) As Boolean
            If Not tokenEnumerator.IsEndOfList AndAlso tokenEnumerator.Current.Token = optionalToken Then
                tokenInfo = tokenEnumerator.Current
                tokenEnumerator.MoveNext()
                Return True
            End If

            tokenInfo = TokenInfo.Illegal
            Return False
        End Function

        Friend Function EatOptionalToken(ByVal tokenEnumerator As TokenEnumerator, ByVal optionalToken As Token) As Boolean
            Dim tokenInfo As TokenInfo
            Return EatOptionalToken(tokenEnumerator, optionalToken, tokenInfo)
        End Function

        Friend Function EatSimpleIdentifier(ByVal tokenEnumerator As TokenEnumerator, <Out> ByRef tokenInfo As TokenInfo) As Boolean
            If Not tokenEnumerator.IsEndOfList AndAlso tokenEnumerator.Current.Token = Token.Identifier Then
                tokenInfo = tokenEnumerator.Current
                tokenEnumerator.MoveNext()
                If tokenInfo.Text <> "_" Then Return True
            End If

            tokenInfo = TokenInfo.Illegal
            AddError(tokenEnumerator.Current, ResourceHelper.GetString("IdentifierExpected"))
            Return False
        End Function


        Friend Function EatExpression(ByVal tokenEnumerator As TokenEnumerator, <Out> ByRef expression As Expression) As Boolean
            expression = BuildArithmeticExpression(tokenEnumerator)

            If expression IsNot Nothing Then Return True

            AddError(tokenEnumerator.Current, ResourceHelper.GetString("ExpressionExpected"))
            Return False
        End Function

        Friend Function EatLogicalExpression(ByVal tokenEnumerator As TokenEnumerator, <Out> ByRef expression As Expression) As Boolean
            expression = BuildLogicalExpression(tokenEnumerator)

            If expression IsNot Nothing Then
                Return True
            End If

            AddError(tokenEnumerator.Current, ResourceHelper.GetString("ConditionExpected"))
            Return False
        End Function

        Friend Function ExpectEndOfLine(ByVal tokenEnumerator As TokenEnumerator) As Boolean
            If tokenEnumerator.IsEndOfNonCommentList Then
                Return True
            End If

            AddError(tokenEnumerator.Current, String.Format(ResourceHelper.GetString("UnexpectedTokenFound"), tokenEnumerator.Current.Text))
            Return False
        End Function

    End Class
End Namespace
