﻿Imports System
Imports System.Collections.Generic
Imports System.IO
Imports Microsoft.SmallBasic.Expressions
Imports Microsoft.SmallBasic.Library
Imports Microsoft.SmallBasic.Statements
Imports System.Runtime.InteropServices

Namespace Microsoft.SmallBasic
    Public Class Parser
        Private _reader As TextReader
        Private _currentLine As Integer
        Private _currentLineEnumerator As TokenEnumerator
        Private _rewindRequested As Boolean

        Public Property Errors As List(Of [Error])

        Public ReadOnly Property ParseTree As List(Of Statement)


        Public ReadOnly Property SymbolTable As SymbolTable

        Private Function ConstructWhileStatement(ByVal tokenEnumerator As TokenEnumerator) As WhileStatement
            Dim whileStatement As WhileStatement = New WhileStatement()
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
                    whileStatement2.EndWhileToken = tokenEnumerator.Current
                    flag = True
                    Exit While
                End If

                If tokenEnumerator.Current.Token = Token.Sub OrElse tokenEnumerator.Current.Token = Token.EndSub Then
                    Exit While
                End If

                whileStatement2.WhileBody.Add(GetStatementFromTokens(tokenEnumerator))
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
            Dim forStatement As ForStatement = New ForStatement()
            forStatement.StartToken = tokenEnumerator.Current
            forStatement.ForToken = tokenEnumerator.Current
            Dim forStatement2 = forStatement
            tokenEnumerator.MoveNext()

            If EatSimpleIdentifier(tokenEnumerator, forStatement2.Iterator) AndAlso EatToken(tokenEnumerator, Token.Equals) AndAlso EatExpression(tokenEnumerator, forStatement2.InitialValue) AndAlso EatToken(tokenEnumerator, Token.To, forStatement2.ToToken) AndAlso EatExpression(tokenEnumerator, forStatement2.FinalValue) AndAlso (Not EatOptionalToken(tokenEnumerator, Token.Step, forStatement2.StepToken) OrElse EatExpression(tokenEnumerator, forStatement2.StepValue)) Then
                ExpectEndOfLine(tokenEnumerator)
            End If

            tokenEnumerator = ReadNextLine()
            Dim flag = False

            While tokenEnumerator IsNot Nothing

                If tokenEnumerator.Current.Token = Token.EndFor OrElse tokenEnumerator.Current.Token = Token.Next Then
                    forStatement2.EndForToken = tokenEnumerator.Current
                    flag = True
                    Exit While
                End If

                If tokenEnumerator.Current.Token = Token.Sub OrElse tokenEnumerator.Current.Token = Token.EndSub Then
                    Exit While
                End If

                forStatement2.ForBody.Add(GetStatementFromTokens(tokenEnumerator))
                tokenEnumerator = ReadNextLine()
            End While

            If flag Then
                tokenEnumerator.MoveNext()
                ExpectEndOfLine(tokenEnumerator)
            Else
                AddError(ResourceHelper.GetString("EndForExpected"))
            End If

            Return forStatement2
        End Function

        Private Function ConstructIfStatement(ByVal tokenEnumerator As TokenEnumerator) As IfStatement
            Dim ifStatement As IfStatement = New IfStatement()
            ifStatement.StartToken = tokenEnumerator.Current
            ifStatement.IfToken = tokenEnumerator.Current
            Dim ifStatement2 = ifStatement
            tokenEnumerator.MoveNext()

            If EatLogicalExpression(tokenEnumerator, ifStatement2.Condition) AndAlso EatToken(tokenEnumerator, Token.Then, ifStatement2.ThenToken) Then
                ExpectEndOfLine(tokenEnumerator)
            End If

            tokenEnumerator = ReadNextLine()
            Dim flag = False
            Dim flag2 = False
            Dim flag3 = False

            While tokenEnumerator IsNot Nothing

                If tokenEnumerator.Current.Token = Token.EndIf Then
                    ifStatement2.EndIfToken = tokenEnumerator.Current
                    flag = True
                    Exit While
                End If

                If tokenEnumerator.Current.Token = Token.Else Then
                    ifStatement2.ElseToken = tokenEnumerator.Current
                    flag3 = True
                    Exit While
                End If

                If tokenEnumerator.Current.Token = Token.ElseIf Then
                    flag2 = True
                    Exit While
                End If

                If tokenEnumerator.Current.Token = Token.Sub OrElse tokenEnumerator.Current.Token = Token.EndSub Then
                    Exit While
                End If

                ifStatement2.ThenStatements.Add(GetStatementFromTokens(tokenEnumerator))
                tokenEnumerator = ReadNextLine()
            End While

            If flag2 Then
                While tokenEnumerator IsNot Nothing

                    If tokenEnumerator.Current.Token = Token.EndIf Then
                        ifStatement2.EndIfToken = tokenEnumerator.Current
                        flag = True
                        Exit While
                    End If

                    If tokenEnumerator.Current.Token = Token.Else Then
                        ifStatement2.ElseToken = tokenEnumerator.Current
                        flag3 = True
                        Exit While
                    End If

                    If tokenEnumerator.Current.Token = Token.Sub OrElse tokenEnumerator.Current.Token = Token.EndSub OrElse tokenEnumerator.Current.Token <> Token.ElseIf Then
                        Exit While
                    End If

                    ifStatement2.ElseIfStatements.Add(ConstructElseIfStatement(tokenEnumerator))
                End While
            End If

            If flag3 Then
                tokenEnumerator.MoveNext()
                ExpectEndOfLine(tokenEnumerator)
                tokenEnumerator = ReadNextLine()

                While tokenEnumerator IsNot Nothing

                    If tokenEnumerator.Current.Token = Token.EndIf Then
                        ifStatement2.EndIfToken = tokenEnumerator.Current
                        flag = True
                        Exit While
                    End If

                    If tokenEnumerator.Current.Token = Token.Sub OrElse tokenEnumerator.Current.Token = Token.EndSub OrElse tokenEnumerator.Current.Token = Token.Else Then
                        Exit While
                    End If

                    ifStatement2.ElseStatements.Add(GetStatementFromTokens(tokenEnumerator))
                    tokenEnumerator = ReadNextLine()
                End While
            End If

            If flag Then
                tokenEnumerator.MoveNext()
                ExpectEndOfLine(tokenEnumerator)
            Else
                AddError(ResourceHelper.GetString("EndIfExpected"))
            End If

            Return ifStatement2
        End Function

        Private Function ConstructElseIfStatement(ByRef tokenEnumerator As TokenEnumerator) As ElseIfStatement
            Dim elseIfStatement As ElseIfStatement = New ElseIfStatement()
            elseIfStatement.StartToken = tokenEnumerator.Current
            elseIfStatement.ElseIfToken = tokenEnumerator.Current
            Dim elseIfStatement2 = elseIfStatement
            tokenEnumerator.MoveNext()

            If EatLogicalExpression(tokenEnumerator, elseIfStatement2.Condition) AndAlso EatToken(tokenEnumerator, Token.Then, elseIfStatement2.ThenToken) Then
                ExpectEndOfLine(tokenEnumerator)
            End If

            tokenEnumerator = ReadNextLine()

            While tokenEnumerator IsNot Nothing AndAlso tokenEnumerator.Current.Token <> Token.EndIf AndAlso tokenEnumerator.Current.Token <> Token.Else AndAlso tokenEnumerator.Current.Token <> Token.ElseIf AndAlso tokenEnumerator.Current.Token <> Token.Sub AndAlso tokenEnumerator.Current.Token <> Token.EndSub
                elseIfStatement2.ThenStatements.Add(GetStatementFromTokens(tokenEnumerator))
                tokenEnumerator = ReadNextLine()
            End While

            Return elseIfStatement2
        End Function

        Private Function ConstructSubStatement(ByVal tokenEnumerator As TokenEnumerator) As SubroutineStatement
            Dim subroutineStatement As SubroutineStatement = New SubroutineStatement()
            subroutineStatement.StartToken = tokenEnumerator.Current
            subroutineStatement.SubToken = tokenEnumerator.Current
            Dim subroutineStatement2 = subroutineStatement
            tokenEnumerator.MoveNext()
            EatSimpleIdentifier(tokenEnumerator, subroutineStatement2.SubroutineName)
            ExpectEndOfLine(tokenEnumerator)
            tokenEnumerator = ReadNextLine()
            Dim flag = False

            While tokenEnumerator IsNot Nothing

                If tokenEnumerator.Current.Token = Token.EndSub Then
                    subroutineStatement2.EndSubToken = tokenEnumerator.Current
                    flag = True
                    Exit While
                End If

                If tokenEnumerator.Current.Token = Token.Sub Then
                    RewindLine()
                    Exit While
                End If

                subroutineStatement2.SubroutineBody.Add(GetStatementFromTokens(tokenEnumerator))
                tokenEnumerator = ReadNextLine()
            End While

            If flag Then
                tokenEnumerator.MoveNext()
                ExpectEndOfLine(tokenEnumerator)
            Else
                AddError(ResourceHelper.GetString("EndSubExpected"))
            End If

            Return subroutineStatement2
        End Function

        Private Function ConstructGotoStatement(ByVal tokenEnumerator As TokenEnumerator) As GotoStatement
            Dim gotoStatement As GotoStatement = New GotoStatement()
            gotoStatement.StartToken = tokenEnumerator.Current
            gotoStatement.GotoToken = tokenEnumerator.Current
            Dim gotoStatement2 = gotoStatement
            tokenEnumerator.MoveNext()
            EatSimpleIdentifier(tokenEnumerator, gotoStatement2.Label)
            ExpectEndOfLine(tokenEnumerator)
            Return gotoStatement2
        End Function

        Private Function ConstructLabelStatement(ByVal tokenEnumerator As TokenEnumerator) As LabelStatement
            Dim labelStatement As LabelStatement = New LabelStatement()
            labelStatement.StartToken = tokenEnumerator.Current
            labelStatement.LabelToken = tokenEnumerator.Current
            Dim labelStatement2 = labelStatement
            tokenEnumerator.MoveNext()

            If EatToken(tokenEnumerator, Token.Colon, labelStatement2.ColonToken) Then
                ExpectEndOfLine(tokenEnumerator)
            End If

            Return labelStatement2
        End Function

        Private Function ConstructIdentifierStatement(ByVal tokenEnumerator As TokenEnumerator) As Statement
            Dim current = tokenEnumerator.Current
            Dim tokenInfo As TokenInfo = tokenEnumerator.PeekNext()

            If tokenInfo.Equals(TokenInfo.Illegal) Then
                AddError(tokenEnumerator.Current, ResourceHelper.GetString("UnrecognizedStatement"))
                Dim leftValue = BuildIdentifierTerm(tokenEnumerator)
                Dim assignmentStatement As AssignmentStatement = New AssignmentStatement()
                assignmentStatement.StartToken = current
                assignmentStatement.LeftValue = leftValue
                Return assignmentStatement
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
                Dim methodCallStatement As MethodCallStatement = New MethodCallStatement()
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

        Private Function ConstructSubroutineCallStatement(ByVal tokenEnumerator As TokenEnumerator) As Statement
            Dim subroutineCallStatement As SubroutineCallStatement = New SubroutineCallStatement()
            subroutineCallStatement.StartToken = tokenEnumerator.Current
            subroutineCallStatement.SubroutineName = tokenEnumerator.Current
            Dim result = subroutineCallStatement
            tokenEnumerator.MoveNext()

            If EatToken(tokenEnumerator, Token.LeftParens) AndAlso EatToken(tokenEnumerator, Token.RightParens) Then
                ExpectEndOfLine(tokenEnumerator)
            End If

            Return result
        End Function

        Public Function BuildArithmeticExpression(tokenEnumerator As TokenEnumerator) As Expression
            Return BuildExpression(tokenEnumerator, includeLogical:=False)
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

        Private Function BuildTerm(ByVal tokenEnumerator As TokenEnumerator, ByVal includeLogical As Boolean) As Expression
            Dim current = tokenEnumerator.Current

            If tokenEnumerator.IsEndOfList OrElse tokenEnumerator.Current.Token = Token.Illegal Then
                Return Nothing
            End If

            Dim expression As Expression
            Select Case tokenEnumerator.Current.Token
                Case Token.StringLiteral, Token.NumericLiteral
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
                        arguments:=ParseCommaSepList(tokenEnumerator, Token.RightCurlyBracket)
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

                    Dim negativeExpression As NegativeExpression = New NegativeExpression()
                    negativeExpression.Negation = tokenEnumerator.Current
                    negativeExpression.Expression = expression2
                    negativeExpression.Precedence = 9
                    expression = negativeExpression
            End Select

            expression.StartToken = current
            expression.EndToken = tokenEnumerator.Current
            Return expression
        End Function

        Private Function MergeExpression(ByVal leftHandExpr As Expression, ByVal rightHandExpr As Expression, ByVal operatorToken As TokenInfo) As Expression
            If TypeOf leftHandExpr Is InitializerExpression OrElse TypeOf rightHandExpr Is InitializerExpression Then
                _Errors.Add(New [Error](operatorToken, "Array initializer can't be used as an operand in binary operations"))
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

            If tokenEnumerator.Current.Token = Token.Dot Then
                tokenEnumerator.MoveNext()

                If Not EatSimpleIdentifier(tokenEnumerator, variable) Then
                    Return Nothing
                End If

                If EatOptionalToken(tokenEnumerator, Token.LeftParens) Then
                    Dim methodCallExpression As New MethodCallExpression(
                        startToken:=current,
                        precedence:=9,
                        typeName:=current,
                        methodName:=variable,
                        arguments:=ParseCommaSepList(tokenEnumerator, Token.RightParens)
                   )

                    methodCallExpression.EndToken = tokenEnumerator.Current
                    tokenEnumerator.MoveNext()
                    Return methodCallExpression
                End If

                Dim propertyExpression As PropertyExpression = New PropertyExpression()
                propertyExpression.StartToken = current
                propertyExpression.EndToken = variable
                propertyExpression.Precedence = 9
                propertyExpression.TypeName = current
                propertyExpression.PropertyName = variable
                Return propertyExpression
            End If

            If tokenEnumerator.Current.Token = Token.LeftBracket Then
                Dim identifierExpression As New IdentifierExpression()
                identifierExpression.StartToken = current
                identifierExpression.Identifier = current
                identifierExpression.EndToken = current
                identifierExpression.Precedence = 9
                Dim leftHand As Expression = identifierExpression
                Dim expression2 As Expression

                While True
                    EatToken(tokenEnumerator, Token.LeftBracket)
                    expression2 = BuildArithmeticExpression(tokenEnumerator)

                    If expression2 Is Nothing Then
                        _Errors.Add(New [Error](tokenEnumerator.Current, ResourceHelper.GetString("ExpressionExpected")))
                        Exit While
                    End If

                    EatToken(tokenEnumerator, Token.RightBracket)

                    If tokenEnumerator.Current.Token <> Token.LeftBracket Then
                        Exit While
                    End If

                    Dim arrayExpression As New ArrayExpression()
                    arrayExpression.StartToken = current
                    arrayExpression.EndToken = tokenEnumerator.Current
                    arrayExpression.Precedence = 9
                    arrayExpression.LeftHand = leftHand
                    arrayExpression.Indexer = expression2
                    leftHand = arrayExpression
                End While

                Dim arrayExpression2 As New ArrayExpression()
                arrayExpression2.StartToken = current
                arrayExpression2.EndToken = tokenEnumerator.Current
                arrayExpression2.Precedence = 9
                arrayExpression2.LeftHand = leftHand
                arrayExpression2.Indexer = expression2
                Return arrayExpression2
            End If

            Dim identifierExpression2 As New IdentifierExpression()
            identifierExpression2.StartToken = current
            identifierExpression2.EndToken = current
            identifierExpression2.Precedence = 9
            identifierExpression2.Identifier = current
            Return identifierExpression2
        End Function

        Private Function ParseCommaSepList(tokenEnumerator As TokenEnumerator, closingToken As Token) As List(Of Expression)
            Dim items As New List(Of Expression)

            While tokenEnumerator.Current.Token <> closingToken
                Dim expression = BuildArithmeticExpression(tokenEnumerator)

                If expression Is Nothing Then
                    _Errors.Add(New [Error](tokenEnumerator.Current, ResourceHelper.GetString("ExpressionExpected")))
                    Exit While
                End If

                items.Add(expression)
                EatOptionalToken(tokenEnumerator, Token.Comma)

                If tokenEnumerator.IsEndOfNonCommentList Then
                    _Errors.Add(New [Error](tokenEnumerator.Current, ResourceHelper.GetString("UnexpectedMethodCallEOL")))
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
            Me.New(Nothing)
        End Sub

        Public Sub New(ByVal errors As List(Of [Error]))
            _Errors = errors

            If _Errors Is Nothing Then
                _Errors = New List(Of [Error])()
            End If

            _SymbolTable = New SymbolTable(_Errors)
        End Sub

        Public Sub Parse(ByVal reader As TextReader)
            If reader Is Nothing Then
                Throw New ArgumentNullException("reader")
            End If

            _Errors.Clear()
            _SymbolTable.Reset()
            _ParseTree = New List(Of Statement)()
            _reader = reader
            BuildParseTree()

            For Each item In _ParseTree
                item.AddSymbols(_SymbolTable)
            Next
        End Sub

        Public Shared Function Parse(code As String) As Parser
            Dim _parser As New Parser
            _parser.Parse(New IO.StringReader(code))
            Return _parser
        End Function

        Public Sub ParseMore(code As String)
            _reader = New IO.StringReader(code)
            BuildParseTree()

            For Each item In _ParseTree
                item.AddSymbols(_SymbolTable)
            Next
        End Sub

        Private Sub BuildParseTree()
            _currentLine = -1

            While True
                Dim tokenEnumerator = ReadNextLine()

                If tokenEnumerator Is Nothing Then
                    Exit While
                End If

                Dim statementFromTokens = GetStatementFromTokens(tokenEnumerator)
                _ParseTree.Add(statementFromTokens)
            End While
        End Sub

        Friend Function ReadNextLine() As TokenEnumerator
            If _rewindRequested Then
                _rewindRequested = False
                Return _currentLineEnumerator
            End If

            Dim text As String = _reader.ReadLine()

            If Equals(text, Nothing) Then
                Return Nothing
            End If

            _currentLine += 1
            _currentLineEnumerator = New LineScanner().GetTokenList(text, _currentLine)
            Return _currentLineEnumerator
        End Function

        Friend Sub RewindLine()
            _rewindRequested = True
        End Sub

        Friend Function GetStatementFromTokens(ByVal tokenEnumerator As TokenEnumerator) As Statement
            If tokenEnumerator.IsEndOfList Then
                Dim statement As Statement = New EmptyStatement()
                statement.StartToken = New TokenInfo With {
                    .Line = tokenEnumerator.LineNumber
                }
                Return statement
            End If

            Dim statement2 As Statement = Nothing

            Select Case tokenEnumerator.Current.Token
                Case Token.While
                    statement2 = ConstructWhileStatement(tokenEnumerator)
                Case Token.For
                    statement2 = ConstructForStatement(tokenEnumerator)
                Case Token.Goto
                    statement2 = ConstructGotoStatement(tokenEnumerator)
                Case Token.If
                    statement2 = ConstructIfStatement(tokenEnumerator)
                Case Token.ElseIf
                    AddError(tokenEnumerator.Current, String.Format(ResourceHelper.GetString("ElseIfUnexpected"), tokenEnumerator.Current.Text))
                    Return New IllegalStatement()
                Case Token.Sub
                    statement2 = ConstructSubStatement(tokenEnumerator)
                Case Token.Identifier
                    statement2 = ConstructIdentifierStatement(tokenEnumerator)
                Case Token.Comment
                    Dim emptyStatement As EmptyStatement = New EmptyStatement()
                    emptyStatement.StartToken = tokenEnumerator.Current
                    statement2 = emptyStatement
                    Exit Select
            End Select

            If statement2 Is Nothing Then
                AddError(tokenEnumerator.Current, String.Format(ResourceHelper.GetString("UnexpectedTokenAtLocation"), tokenEnumerator.Current.Text))
                Return New IllegalStatement()
            End If

            If tokenEnumerator.Current.Token = Token.Comment Then
                statement2.EndingComment = tokenEnumerator.Current
            End If

            Return statement2
        End Function

        Friend Sub AddError(ByVal errorDescription As String)
            AddError(_currentLine, errorDescription)
        End Sub

        Friend Sub AddError(ByVal line As Integer, ByVal errorDescription As String)
            AddError(line, 0, errorDescription)
        End Sub

        Friend Sub AddError(ByVal line As Integer, ByVal column As Integer, ByVal errorDescription As String)
            _Errors.Add(New [Error](line, column, errorDescription))
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
                        Case Token.Dot, Token.LeftParens, Token.RightParens, Token.LeftBracket, Token.RightBracket
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
            _Errors.Add(New [Error](tokenEnumerator.Current, String.Format(ResourceHelper.GetString("TokenExpected"), expectedToken)))
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

        Friend Function EatSimpleIdentifier(ByVal tokenEnumerator As TokenEnumerator, <Out> ByRef variable As TokenInfo) As Boolean
            If Not tokenEnumerator.IsEndOfList AndAlso tokenEnumerator.Current.Token = Token.Identifier Then
                variable = tokenEnumerator.Current
                tokenEnumerator.MoveNext()
                Return True
            End If

            variable = TokenInfo.Illegal
            _Errors.Add(New [Error](tokenEnumerator.Current, ResourceHelper.GetString("IdentifierExpected")))
            Return False
        End Function

        Friend Function EatExpression(ByVal tokenEnumerator As TokenEnumerator, <Out> ByRef expression As Expression) As Boolean
            expression = BuildArithmeticExpression(tokenEnumerator)

            If expression IsNot Nothing Then
                Return True
            End If

            _Errors.Add(New [Error](tokenEnumerator.Current, ResourceHelper.GetString("ExpressionExpected")))
            Return False
        End Function

        Friend Function EatLogicalExpression(ByVal tokenEnumerator As TokenEnumerator, <Out> ByRef expression As Expression) As Boolean
            expression = BuildLogicalExpression(tokenEnumerator)

            If expression IsNot Nothing Then
                Return True
            End If

            _Errors.Add(New [Error](tokenEnumerator.Current, ResourceHelper.GetString("ConditionExpected")))
            Return False
        End Function

        Friend Function ExpectEndOfLine(ByVal tokenEnumerator As TokenEnumerator) As Boolean
            If tokenEnumerator.IsEndOfNonCommentList Then
                Return True
            End If

            _Errors.Add(New [Error](tokenEnumerator.Current, String.Format(ResourceHelper.GetString("UnexpectedTokenFound"), tokenEnumerator.Current.Text)))
            Return False
        End Function

    End Class
End Namespace
