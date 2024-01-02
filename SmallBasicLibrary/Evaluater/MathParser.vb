Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Globalization
Imports Microsoft.SmallVisualBasic.Evaluator.Expressions

Namespace Evaluator
    Friend Class MathParser
        Private MathExpression As String
        Private _currentLineEnum As TokenEnumerator

        Public Property Errors As List(Of [Error])

        Private Function BuildExpression(tokenEnum As TokenEnumerator, Optional firstOperationOnly As Boolean = False) As Expression
            If tokenEnum.IsEnd Then
                Return Nothing
            End If

            Dim current = tokenEnum.Current

            Dim leftHandExpr = BuildTerm(tokenEnum)
            If leftHandExpr Is Nothing Then Return Nothing

            While IsValidOperator(tokenEnum.Current.Type)
                Dim current2 = tokenEnum.Current
                tokenEnum.MoveNext()
                Dim rightHandExpr = BuildTerm(tokenEnum)
                If rightHandExpr Is Nothing Then Return Nothing

                leftHandExpr = MergeExpression(leftHandExpr, rightHandExpr, current2)
                If leftHandExpr Is Nothing Then Return Nothing
                If firstOperationOnly Then Exit While
            End While

            leftHandExpr.StartToken = current
            leftHandExpr.EndToken = tokenEnum.Current
            Return leftHandExpr
        End Function

        Private Function BuildTerm(tokenEnum As TokenEnumerator, Optional takePower As Boolean = False) As Expression
            Dim current = tokenEnum.Current

            If tokenEnum.IsEnd OrElse tokenEnum.Current.Type = TokenType.Illegal Then
                Return Nothing
            End If

            Dim expression As Expression
            Select Case tokenEnum.Current.Type
                Case TokenType.NumericLiteral
                    expression = New LiteralExpression(tokenEnum.Current)
                    expression.Precedence = 9
                    tokenEnum.MoveNext()

                Case TokenType.LeftParens
                    tokenEnum.MoveNext()
                    expression = BuildExpression(tokenEnum)
                    If expression Is Nothing Then Return Nothing
                    expression.Precedence = 9
                    If Not EatToken(tokenEnum, TokenType.RightParens) Then Return Nothing

                Case TokenType.Identifier
                    expression = BuildIdentifierTerm(tokenEnum)
                    If expression Is Nothing Then Return Nothing

                Case Else
                    If tokenEnum.Current.Type <> TokenType.Subtraction Then
                        Return Nothing
                    End If

                    tokenEnum.MoveNext()
                    Dim expression2 = BuildTerm(tokenEnum, True)
                    If expression2 Is Nothing Then Return Nothing

                    expression = New NegativeExpression() With {
                        .Negation = tokenEnum.Current,
                        .Expression = expression2,
                        .Precedence = 9
                    }
            End Select

            If takePower Then
                Dim op = tokenEnum.Current
                If op.Type = TokenType.Power Then
                    tokenEnum.MoveNext()
                    Dim rightHand = BuildTerm(tokenEnum)
                    expression = MergeExpression(expression, rightHand, op)
                End If
            End If

            expression.StartToken = current
            expression.EndToken = tokenEnum.Current
            Return expression
        End Function

        Private Function MergeExpression(leftHandExpr As Expression, rightHandExpr As Expression, operatorToken As Token) As Expression
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
            Dim tokenType = tokenEnum.Current.Type

            If EatOptionalToken(tokenEnum, TokenType.LeftParens) Then
                Dim closeTokenFound = False
                Dim methodCallExpression As New MethodCallExpression(
                        startToken:=current,
                        precedence:=9,
                        methodName:=current,
                        arguments:=ParseCommaSepaeatedList(tokenEnum, TokenType.RightParens, closeTokenFound, False)
                )

                methodCallExpression.EndToken = tokenEnum.Current
                If closeTokenFound Then tokenEnum.MoveNext()
                Return methodCallExpression
            End If

            Return New IdentifierExpression() With {
                .StartToken = current,
                .EndToken = current,
                .Precedence = 9,
                .Identifier = current
            }
        End Function

        Private Function ParseCommaSepaeatedList(
                           tokenEnum As TokenEnumerator,
                           closeToken As TokenType,
                           <Out> ByRef closeTokenFound As Boolean,
                           Optional commaIsOptional As Boolean = True
                  ) As List(Of Expression)

            Dim items As New List(Of Expression)
            closeTokenFound = False
            Dim exprExpected = False

            Do
                Dim current = tokenEnum.Current
                If current.Type = closeToken Then
                    closeTokenFound = True
                    If exprExpected Then
                        Dim exprToken As New Token With {
                            .Line = current.Line,
                            .Column = If(tokenEnum.IsEnd, current.EndColumn, current.Column)
                        }
                        AddError(exprToken, "Expression is expected.")
                    End If

                    Exit Do
                End If

                Dim expression = BuildExpression(tokenEnum)

                If expression Is Nothing Then
                    AddError(tokenEnum.Current, "Expression is expected")
                    Exit Do
                End If

                items.Add(expression)
                exprExpected = False

                If tokenEnum.IsEnd Then Exit Do

                Dim comma = tokenEnum.Current
                If EatOptionalToken(tokenEnum, TokenType.Comma) Then
                    exprExpected = True
                ElseIf Not commaIsOptional Then
                    current = tokenEnum.Current
                    Select Case current.Type
                        Case closeToken
                            closeTokenFound = True

                        Case TokenType.RightParens
                            ' outer list is closed, and the inner list misses the closeToken
                        Case Else
                            TokenIsExpected(tokenEnum, TokenType.Comma)
                    End Select

                    Exit Do
                End If

                If tokenEnum.IsEnd Then
                    If closeTokenFound Then
                        AddError(tokenEnum.Current, "Unexpected end of line")
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
                Case TokenType.RightParens
                    expectedChar = ")"
                Case TokenType.Comma
                    expectedChar = ","
            End Select
            AddError(token, $"`{expectedChar}` is expected here but not found")
        End Sub

        Private Function IsValidOperator(tokenType As TokenType) As Boolean
            Dim operatorPriority = GetOperatorPriority(tokenType)
            If operatorPriority = 0 Then
                Dim t = _currentLineEnum.Current
                AddError(t, $"Unrecognized operator {t.Text}")
            End If

            Return operatorPriority >= 5
        End Function

        Friend Shared Function GetOperatorPriority(token As TokenType) As Integer
            Select Case token
                Case TokenType.Power
                    Return 9
                Case TokenType.Division, TokenType.Multiplication
                    Return 7
                Case TokenType.Mod
                    Return 6
                Case TokenType.Addition, TokenType.Subtraction
                    Return 5
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
        End Sub

        Public Function Parse(expr As String) As Expression
            Me.MathExpression = expr
            _Errors.Clear()
            Dim line = MathExpression.Replace(vbCr, " ").Replace(vbLf, " ")
            _currentLineEnum = LineScanner.GetTokenEnumerator(line)
            Return BuildExpression(_currentLineEnum)
        End Function

        Friend Sub AddError(token As Token, errorDescription As String)
            _Errors.Add(New [Error](token, errorDescription))
        End Sub

        Friend Function EatToken(tokenEnum As TokenEnumerator, expectedToken As TokenType, <Out> ByRef token As Token) As Boolean
            If Not tokenEnum.IsEnd AndAlso tokenEnum.Current.Type = expectedToken Then
                token = tokenEnum.Current
                tokenEnum.MoveNext()
                Return True
            End If

            token = Token.Illegal
            AddError(tokenEnum.Current, String.Format("A {0} token is expected here", expectedToken))
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

    End Class

    Public Class [Error]
        Public ReadOnly Property Line As Integer
        Public ReadOnly Property Column As Integer
        Public ReadOnly Property Description As String

        Public ReadOnly Property Token As Token

        Public Sub New(token As Token, description As String)
            Me.New(token.Line, token.Column, description)
            _Token = token
        End Sub

        Public Sub New(line As Integer, column As Integer, description As String)
            _Line = line
            _Column = column
            _Description = description
        End Sub

        Public Overrides Function ToString() As String
            Return $"{Line},{Column}: {Description}"
        End Function
    End Class

End Namespace
