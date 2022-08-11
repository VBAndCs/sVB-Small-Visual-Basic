Imports System.Collections.Generic
Imports System.Globalization
Imports System.Reflection.Emit
Imports System.Text
Imports Microsoft.SmallBasic.Completion
Imports Microsoft.SmallBasic.Expressions
Imports Microsoft.SmallBasic.Library

Namespace Microsoft.SmallBasic.Statements
    Public Class ForStatement
        Inherits LoopStatement

        Public InitialValue As Expression
        Public FinalValue As Expression
        Public StepValue As Expression
        Public Iterator As TokenInfo
        Public ForToken As TokenInfo
        Public ToToken As TokenInfo
        Public StepToken As TokenInfo
        Public Subroutine As SubroutineStatement
        Public EqualsToken As TokenInfo

        Public Overrides Function GetStatement(lineNumber) As Statement
            If lineNumber < StartToken.Line Then Return Nothing
            If lineNumber > EndLoopToken.Line Then Return Nothing
            If lineNumber <= Iterator.Line Then Return Me
            If lineNumber <= EqualsToken.Line Then Return Me
            If lineNumber <= InitialValue?.EndToken.Line Then Return Me
            If lineNumber <= ToToken.Line Then Return Me
            If lineNumber <= FinalValue?.EndToken.Line Then Return Me
            If lineNumber <= StepToken.Line Then Return Me
            If lineNumber <= StepValue?.EndToken.Line Then Return Me

            For Each statment In Body
                Dim st = statment.GetStatement(lineNumber)
                If st IsNot Nothing Then Return st
            Next

            If lineNumber <= EndLoopToken.Line Then Return Me

            Return Nothing
        End Function

        Public Overrides Function GetKeywords() As LegalTokens
            Dim spans As New LegalTokens
            spans.Add(ForToken)
            spans.Add(ToToken)
            spans.Add(StepToken)
            spans.Add(EndLoopToken)
            Return spans
        End Function

        Public Overrides Sub AddSymbols(ByVal symbolTable As SymbolTable)
            MyBase.AddSymbols(symbolTable)
            Iterator.Parent = Me
            ForToken.Parent = Me
            ToToken.Parent = Me
            StepToken.Parent = Me
            EndLoopToken.Parent = Me
            EqualsToken.Parent = Me

            If Iterator.Token <> Token.Illegal Then
                Dim id As New IdentifierExpression() With {
                    .Identifier = Iterator,
                    .Subroutine = Subroutine
                }
                symbolTable.AddVariable(id, True)
            End If

            If InitialValue IsNot Nothing Then
                InitialValue.Parent = Me
                InitialValue.AddSymbols(symbolTable)
            End If

            If FinalValue IsNot Nothing Then
                FinalValue.Parent = Me
                FinalValue.AddSymbols(symbolTable)
            End If

            If StepValue IsNot Nothing Then
                StepValue.Parent = Me
                StepValue.AddSymbols(symbolTable)
            End If

            For Each item In Body
                item.Parent = Me
                item.AddSymbols(symbolTable)
            Next
        End Sub

        Public Overrides Sub PrepareForEmit(ByVal scope As CodeGenScope)
            For Each item In Body
                item.PrepareForEmit(scope)
            Next
        End Sub

        Public Overrides Sub EmitIL(ByVal scope As CodeGenScope)
            Dim iteratorVar = scope.CreateLocalBuilder(Subroutine, Iterator)

            InitialValue.EmitIL(scope)
            scope.ILGenerator.Emit(OpCodes.Stloc, iteratorVar)

            ContinueLabel = scope.ILGenerator.DefineLabel()
            ExitLabel = scope.ILGenerator.DefineLabel()
            Dim ConditionLabel = scope.ILGenerator.DefineLabel()
            Dim lblElseIf = scope.ILGenerator.DefineLabel()
            Dim loopBody = scope.ILGenerator.DefineLabel()

            scope.ILGenerator.MarkLabel(ConditionLabel)
            scope.ILGenerator.Emit(OpCodes.Ldloc, iteratorVar)
            FinalValue.EmitIL(scope)

            If StepValue IsNot Nothing Then
                StepValue.EmitIL(scope)
                scope.ILGenerator.Emit(OpCodes.Ldc_R8, 0.0)
                scope.ILGenerator.EmitCall(OpCodes.Call, scope.TypeInfoBag.NumberToPrimitive, Nothing)
                scope.ILGenerator.EmitCall(OpCodes.Call, scope.TypeInfoBag.LessThan, Nothing)
                scope.ILGenerator.EmitCall(OpCodes.Call, scope.TypeInfoBag.PrimitiveToBoolean, Nothing)
                scope.ILGenerator.Emit(OpCodes.Brfalse, lblElseIf)
                scope.ILGenerator.EmitCall(OpCodes.Call, scope.TypeInfoBag.GreaterThanOrEqualTo, Nothing)
                scope.ILGenerator.EmitCall(OpCodes.Call, scope.TypeInfoBag.PrimitiveToBoolean, Nothing)
                scope.ILGenerator.Emit(OpCodes.Brfalse, ExitLabel)
                scope.ILGenerator.Emit(OpCodes.Br, loopBody)
            End If

            scope.ILGenerator.MarkLabel(lblElseIf)
            scope.ILGenerator.EmitCall(OpCodes.Call, scope.TypeInfoBag.LessThanOrEqualTo, Nothing)
            scope.ILGenerator.EmitCall(OpCodes.Call, scope.TypeInfoBag.PrimitiveToBoolean, Nothing)
            scope.ILGenerator.Emit(OpCodes.Brfalse, ExitLabel)
            scope.ILGenerator.MarkLabel(loopBody)

            For Each item In Body
                item.EmitIL(scope)
            Next

            ' Increase Iterator
            scope.ILGenerator.MarkLabel(ContinueLabel)
            scope.ILGenerator.Emit(OpCodes.Ldloc, iteratorVar)

            If StepValue IsNot Nothing Then
                StepValue.EmitIL(scope)
            Else
                scope.ILGenerator.Emit(OpCodes.Ldc_R8, 1.0)
                scope.ILGenerator.EmitCall(OpCodes.Call, scope.TypeInfoBag.NumberToPrimitive, Nothing)
            End If

            scope.ILGenerator.EmitCall(OpCodes.Call, scope.TypeInfoBag.Add, Nothing)
            scope.ILGenerator.Emit(OpCodes.Stloc, iteratorVar)
            scope.ILGenerator.Emit(OpCodes.Br, ConditionLabel)
            scope.ILGenerator.MarkLabel(ExitLabel)
        End Sub

        Public Overrides Sub PopulateCompletionItems(
                                 completionBag As CompletionBag,
                                 line As Integer,
                                 column As Integer,
                                 globalScope As Boolean)

            If StartToken.Line = line Then
                If column <= ForToken.EndColumn Then
                    CompletionHelper.FillAllGlobalItems(completionBag, globalScope)

                ElseIf EqualsToken.Token = Token.Illegal OrElse column < EqualsToken.Column Then
                    CompletionHelper.FillLocals(completionBag, Subroutine?.Name.NormalizedText)

                ElseIf ToToken.Token = Token.Illegal OrElse column < ToToken.EndColumn Then
                    CompletionHelper.FillKeywords(completionBag, Token.To)
                    CompletionHelper.FillExpressionItems(completionBag)

                ElseIf StepToken.Token = Token.Illegal OrElse column < StepToken.EndColumn Then
                    CompletionHelper.FillKeywords(completionBag, Token.Step)
                    CompletionHelper.FillExpressionItems(completionBag)

                Else
                    CompletionHelper.FillExpressionItems(completionBag)
                End If

            Else
                Dim statementContaining = GetStatementContaining(Body, line)
                CompletionHelper.FillKeywords(completionBag, Token.EndFor, Token.Next)
                statementContaining?.PopulateCompletionItems(completionBag, line, column, globalScope:=False)
            End If
        End Sub

        Public Overrides Function ToString() As String
            Dim stringBuilder As StringBuilder = New StringBuilder()
            stringBuilder.AppendFormat(CultureInfo.CurrentUICulture, "{0} {1} = {2} To {3}", ForToken.Text, Iterator.Text, InitialValue, FinalValue)

            If StepToken.TokenType <> 0 Then
                stringBuilder.AppendFormat(CultureInfo.CurrentUICulture, "{0} {1}", New Object(1) {StepToken.Text, StepValue})
            End If

            stringBuilder.AppendLine()

            For Each item In Body
                stringBuilder.AppendFormat(CultureInfo.CurrentUICulture, "  {0}", New Object(0) {item})
            Next

            stringBuilder.AppendFormat(CultureInfo.CurrentUICulture, "{0}" & VisualBasic.Constants.vbCrLf, New Object(0) {EndLoopToken.Text})
            Return stringBuilder.ToString()
        End Function
    End Class
End Namespace
