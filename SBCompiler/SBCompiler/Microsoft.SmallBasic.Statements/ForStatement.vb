Imports System.Collections.Generic
Imports System.Globalization
Imports System.Reflection.Emit
Imports System.Text
Imports Microsoft.SmallBasic.Completion
Imports Microsoft.SmallBasic.Expressions
Imports Microsoft.SmallBasic.Library

Namespace Microsoft.SmallBasic.Statements
    Public Class ForStatement
        Inherits Statement

        Public ForBody As List(Of Statement) = New List(Of Statement)()
        Public InitialValue As Expression
        Public FinalValue As Expression
        Public StepValue As Expression
        Public Iterator As TokenInfo
        Public ForToken As TokenInfo
        Public ToToken As TokenInfo
        Public StepToken As TokenInfo
        Public EndForToken As TokenInfo
        Public Subroutine As SubroutineStatement

        Public Overrides Sub AddSymbols(ByVal symbolTable As SymbolTable)
            If InitialValue IsNot Nothing Then
                InitialValue.AddSymbols(symbolTable)
            End If

            If FinalValue IsNot Nothing Then
                FinalValue.AddSymbols(symbolTable)
            End If

            If StepValue IsNot Nothing Then
                StepValue.AddSymbols(symbolTable)
            End If

            For Each item In ForBody
                item.AddSymbols(symbolTable)
            Next
        End Sub

        Public Overrides Sub PrepareForEmit(ByVal scope As CodeGenScope)
            For Each item In ForBody
                item.PrepareForEmit(scope)
            Next
        End Sub

        Public Sub EmitIL2(ByVal scope As CodeGenScope)
            Dim iteratorVar = scope.ILGenerator.DeclareLocal(GetType(Double))
            iteratorVar.SetLocalSymInfo(Iterator.NormalizedText)

            Dim loopBody = scope.ILGenerator.DefineLabel()
            Dim loopCondition = scope.ILGenerator.DefineLabel()
            Dim loopIterator = scope.ILGenerator.DefineLabel()

            InitialValue.EmitIL(scope)
            scope.ILGenerator.Emit(OpCodes.Stloc, iteratorVar)

            scope.ILGenerator.MarkLabel(loopBody)
            scope.ILGenerator.Emit(OpCodes.Ldloc, iteratorVar)
            FinalValue.EmitIL(scope)
            scope.ILGenerator.EmitCall(OpCodes.Call, scope.TypeInfoBag.LessThanOrEqualTo, Nothing)
            scope.ILGenerator.EmitCall(OpCodes.Call, scope.TypeInfoBag.PrimitiveToBoolean, Nothing)
            scope.ILGenerator.Emit(OpCodes.Brfalse, loopIterator)

            Dim iteratorKey = ""
            If Subroutine Is Nothing Then
                iteratorKey = Iterator.NormalizedText
            Else
                iteratorKey = $"{Subroutine.Name.NormalizedText}.{Iterator.NormalizedText}"
            End If

            scope.Locals(iteratorKey) = iteratorVar
            For Each item In ForBody
                item.EmitIL(scope)
            Next
            scope.Locals.Remove(iteratorKey)

            scope.ILGenerator.MarkLabel(loopCondition)

            scope.ILGenerator.Emit(OpCodes.Ldloc, iteratorVar)

            If StepValue IsNot Nothing Then
                StepValue.EmitIL(scope)
            Else
                scope.ILGenerator.Emit(OpCodes.Ldc_R8, 1.0)
                scope.ILGenerator.EmitCall(OpCodes.Call, scope.TypeInfoBag.NumberToPrimitive, Nothing)
            End If

            scope.ILGenerator.EmitCall(OpCodes.Call, scope.TypeInfoBag.Add, Nothing)
            scope.ILGenerator.Emit(OpCodes.Stloc, iteratorVar)
            scope.ILGenerator.Emit(OpCodes.Br, loopBody)
            scope.ILGenerator.MarkLabel(loopIterator)
        End Sub

        Public Overrides Sub EmitIL(ByVal scope As CodeGenScope)
            Dim iteratorVar = scope.ILGenerator.DeclareLocal(GetType(Primitive))
            iteratorVar.SetLocalSymInfo(Iterator.NormalizedText)

            InitialValue.EmitIL(scope)
            scope.ILGenerator.Emit(OpCodes.Stloc, iteratorVar)

            Dim lblExitLoop As Label = scope.ILGenerator.DefineLabel()
            Dim lblLoopIf As Label = scope.ILGenerator.DefineLabel()
            Dim loopBody As Label = scope.ILGenerator.DefineLabel()
            Dim lblEndIf As Label = scope.ILGenerator.DefineLabel()

            scope.ILGenerator.MarkLabel(lblLoopIf)
            scope.ILGenerator.Emit(OpCodes.Ldloc, iteratorVar)
            FinalValue.EmitIL(scope)

            If StepValue IsNot Nothing Then
                StepValue.EmitIL(scope)
                scope.ILGenerator.Emit(OpCodes.Ldc_R8, 0.0)
                scope.ILGenerator.EmitCall(OpCodes.Call, scope.TypeInfoBag.NumberToPrimitive, Nothing)
                scope.ILGenerator.EmitCall(OpCodes.Call, scope.TypeInfoBag.LessThan, Nothing)
                scope.ILGenerator.EmitCall(OpCodes.Call, scope.TypeInfoBag.PrimitiveToBoolean, Nothing)
                scope.ILGenerator.Emit(OpCodes.Brfalse, lblEndIf)
                scope.ILGenerator.EmitCall(OpCodes.Call, scope.TypeInfoBag.GreaterThanOrEqualTo, Nothing)
                scope.ILGenerator.EmitCall(OpCodes.Call, scope.TypeInfoBag.PrimitiveToBoolean, Nothing)
                scope.ILGenerator.Emit(OpCodes.Brfalse, lblExitLoop)
                scope.ILGenerator.Emit(OpCodes.Br, loopBody)
            End If

            scope.ILGenerator.MarkLabel(lblEndIf)
            scope.ILGenerator.EmitCall(OpCodes.Call, scope.TypeInfoBag.LessThanOrEqualTo, Nothing)
            scope.ILGenerator.EmitCall(OpCodes.Call, scope.TypeInfoBag.PrimitiveToBoolean, Nothing)
            scope.ILGenerator.Emit(OpCodes.Brfalse, lblExitLoop)
            scope.ILGenerator.MarkLabel(loopBody)

            Dim iteratorKey = ""
            If Subroutine Is Nothing Then
                iteratorKey = Iterator.NormalizedText
            Else
                iteratorKey = $"{Subroutine.Name.NormalizedText}.{Iterator.NormalizedText}"
            End If

            scope.Locals(iteratorKey) = iteratorVar
            For Each item In ForBody
                item.EmitIL(scope)
            Next

            scope.ILGenerator.Emit(OpCodes.Ldloc, iteratorVar)

            If StepValue IsNot Nothing Then
                StepValue.EmitIL(scope)
            Else
                scope.ILGenerator.Emit(OpCodes.Ldc_R8, 1.0)
                scope.ILGenerator.EmitCall(OpCodes.Call, scope.TypeInfoBag.NumberToPrimitive, Nothing)
            End If

            scope.ILGenerator.EmitCall(OpCodes.Call, scope.TypeInfoBag.Add, Nothing)
            scope.ILGenerator.Emit(OpCodes.Stloc, iteratorVar)
            scope.ILGenerator.Emit(OpCodes.Br, lblLoopIf)
            scope.ILGenerator.MarkLabel(lblExitLoop)
        End Sub

        Public Overrides Sub PopulateCompletionItems(ByVal completionBag As CompletionBag, ByVal line As Integer, ByVal column As Integer, ByVal globalScope As Boolean)
            If StartToken.Line = line Then
                If column <= ForToken.EndColumn Then
                    CompletionHelper.FillAllGlobalItems(completionBag, globalScope)
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
                Dim statementContaining = GetStatementContaining(ForBody, line)
                CompletionHelper.FillKeywords(completionBag, Token.EndFor, Token.Next)
                statementContaining?.PopulateCompletionItems(completionBag, line, column, globalScope:=False)
            End If
        End Sub

        Public Overrides Function GetIndentationLevel(ByVal line As Integer) As Integer
            If line = ForToken.Line Then
                Return 0
            End If

            If EndForToken.Token <> 0 AndAlso line = EndForToken.Line Then
                Return 0
            End If

            Dim statementContaining = GetStatementContaining(ForBody, line)

            If statementContaining IsNot Nothing Then
                Return 1 + statementContaining.GetIndentationLevel(line)
            End If

            Return 1
        End Function

        Public Overrides Function ToString() As String
            Dim stringBuilder As StringBuilder = New StringBuilder()
            stringBuilder.AppendFormat(CultureInfo.CurrentUICulture, "{0} {1} = {2} To {3}", ForToken.Text, Iterator.Text, InitialValue, FinalValue)

            If StepToken.TokenType <> 0 Then
                stringBuilder.AppendFormat(CultureInfo.CurrentUICulture, "{0} {1}", New Object(1) {StepToken.Text, StepValue})
            End If

            stringBuilder.AppendLine()

            For Each item In ForBody
                stringBuilder.AppendFormat(CultureInfo.CurrentUICulture, "  {0}", New Object(0) {item})
            Next

            stringBuilder.AppendFormat(CultureInfo.CurrentUICulture, "{0}" & VisualBasic.Constants.vbCrLf, New Object(0) {EndForToken.Text})
            Return stringBuilder.ToString()
        End Function
    End Class
End Namespace
