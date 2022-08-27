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
        Public Iterator As Token
        Public ForToken As Token
        Public ToToken As Token
        Public StepToken As Token
        Public Subroutine As SubroutineStatement
        Public EqualsToken As Token

        Public Overrides Function GetStatementAt(lineNumber As Integer) As Statement
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
                Dim st = statment.GetStatementAt(lineNumber)
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
            For Each statement In JumpLoopStatements
                spans.Add(statement.StartToken)
            Next
            spans.Add(EndLoopToken)
            Return spans
        End Function

        Public Overrides Sub AddSymbols(symbolTable As SymbolTable)
            MyBase.AddSymbols(symbolTable)
            Iterator.Parent = Me
            ForToken.Parent = Me
            ToToken.Parent = Me
            StepToken.Parent = Me
            EndLoopToken.Parent = Me
            EqualsToken.Parent = Me

            If Iterator.Type <> TokenType.Illegal Then
                Dim id As New IdentifierExpression() With {
                    .Identifier = Iterator,
                    .Subroutine = Subroutine
                }
                Iterator.SymbolType = If(Subroutine Is Nothing, CompletionItemType.GlobalVariable, CompletionItemType.LocalVariable)
                symbolTable.AllIdentifiers.Add(Iterator)
                symbolTable.AddVariable(id, "The for loop counter (iterator)", Subroutine IsNot Nothing)
                If symbolTable.IsGlobalVar(id) Then
                    id.AddSymbolInitialization(symbolTable)
                End If
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

        Public Overrides Sub PrepareForEmit(scope As CodeGenScope)
            For Each item In Body
                item.PrepareForEmit(scope)
            Next
        End Sub

        Public Overrides Sub EmitIL(scope As CodeGenScope)

            InitialValue.EmitIL(scope)

            Dim IsField = Subroutine Is Nothing
            Dim iteratorVar As LocalBuilder
            Dim iteratorField As System.Reflection.FieldInfo

            Dim ConditionLabel = scope.ILGenerator.DefineLabel()
            Dim lblElseIf = scope.ILGenerator.DefineLabel()
            Dim loopBody = scope.ILGenerator.DefineLabel()
            ContinueLabel = scope.ILGenerator.DefineLabel()
            ExitLabel = scope.ILGenerator.DefineLabel()

            If IsField Then
                iteratorField = scope.Fields(Iterator.NormalizedText)
                scope.ILGenerator.Emit(OpCodes.Stsfld, iteratorField)
                scope.ILGenerator.MarkLabel(ConditionLabel)
                scope.ILGenerator.Emit(OpCodes.Ldsfld, iteratorField)
            Else
                iteratorVar = scope.CreateLocalBuilder(Subroutine, Iterator)
                scope.ILGenerator.Emit(OpCodes.Stloc, iteratorVar)
                scope.ILGenerator.MarkLabel(ConditionLabel)
                scope.ILGenerator.Emit(OpCodes.Ldloc, iteratorVar)
            End If

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
            If IsField Then
                scope.ILGenerator.Emit(OpCodes.Ldsfld, iteratorField)
            Else
                scope.ILGenerator.Emit(OpCodes.Ldloc, iteratorVar)
            End If

            If StepValue IsNot Nothing Then
                StepValue.EmitIL(scope)
            Else
                scope.ILGenerator.Emit(OpCodes.Ldc_R8, 1.0)
                scope.ILGenerator.EmitCall(OpCodes.Call, scope.TypeInfoBag.NumberToPrimitive, Nothing)
            End If

            scope.ILGenerator.EmitCall(OpCodes.Call, scope.TypeInfoBag.Add, Nothing)

            If IsField Then
                scope.ILGenerator.Emit(OpCodes.Stsfld, iteratorField)
            Else
                scope.ILGenerator.Emit(OpCodes.Stloc, iteratorVar)
            End If

            scope.ILGenerator.Emit(OpCodes.Br, ConditionLabel)
            scope.ILGenerator.MarkLabel(ExitLabel)
        End Sub

        Public Overrides Sub PopulateCompletionItems(
                                 completionBag As CompletionBag,
                                 line As Integer,
                                 column As Integer,
                                 globalScope As Boolean)

            If ForToken.Line = line AndAlso column <= ForToken.EndColumn Then
                CompletionHelper.FillAllGlobalItems(completionBag, globalScope)

            ElseIf EqualsToken.IsIllegal OrElse line < EqualsToken.Line OrElse
                        (line = EqualsToken.Line AndAlso column < EqualsToken.Column) Then
                CompletionHelper.FillLocals(completionBag, Subroutine?.Name.NormalizedText)
                If Not Iterator.IsIllegal AndAlso Subroutine Is Nothing Then
                    completionBag.CompletionItems.Add(New CompletionItem() With {
                        .Key = Iterator.Text,
                        .DisplayName = Iterator.Text,
                        .ItemType = CompletionItemType.GlobalVariable,
                        .DefinitionIdintifier = completionBag.SymbolTable.GlobalVariables(Iterator.NormalizedText)
                     })
                End If

            ElseIf InForHeader(line) Then
                CompletionHelper.FillKeywords(completionBag, TokenType.To, TokenType.Step)
                CompletionHelper.FillExpressionItems(completionBag)
                CompletionHelper.FillSubroutines(completionBag, True)

            ElseIf Not EndLoopToken.IsIllegal AndAlso line = EndLoopToken.Line Then
                completionBag.CompletionItems.Clear()
                CompletionHelper.FillKeywords(completionBag, EndLoopToken.Type)

            Else
                Dim statement = GetStatementContaining(Body, line)
                CompletionHelper.FillKeywords(completionBag, TokenType.EndFor, TokenType.Next)
                statement?.PopulateCompletionItems(completionBag, line, column, globalScope:=False)
            End If
        End Sub

        Private Function InForHeader(line As Integer) As Boolean
            Dim line2 As Integer

            ' Note that for loop parts can be split over multi lines using the _ symbol,
            ' and some of these parts can be missing while the user is typing code

            If StepValue IsNot Nothing Then
                line2 = StepValue.EndToken.Line

            ElseIf Not StepToken.IsIllegal Then
                line2 = StepToken.Line

            ElseIf FinalValue IsNot Nothing Then
                line2 = FinalValue.EndToken.Line

            ElseIf Not ToToken.IsIllegal Then
                line2 = ToToken.Line

            ElseIf InitialValue IsNot Nothing Then
                line2 = InitialValue.EndToken.Line

            Else
                line2 = EqualsToken.Line
            End If

            Return line <= line2
        End Function

        Public Overrides Function ToString() As String
            Dim stringBuilder As StringBuilder = New StringBuilder()
            stringBuilder.AppendFormat(CultureInfo.CurrentUICulture, "{0} {1} = {2} To {3}", ForToken.Text, Iterator.Text, InitialValue, FinalValue)

            If StepToken.ParseType <> 0 Then
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
