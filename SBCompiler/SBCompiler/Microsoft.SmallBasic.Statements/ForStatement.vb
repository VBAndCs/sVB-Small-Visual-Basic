Imports System.Collections.Generic
Imports System.Globalization
Imports System.Reflection.Emit
Imports System.Text
Imports Microsoft.SmallBasic.Completion
Imports Microsoft.SmallBasic.Expressions

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

        Public Overrides Sub AddSymbols(ByVal symbolTable As SymbolTable)
            If Iterator.Token <> 0 Then
                symbolTable.AddVariable(Iterator)
                symbolTable.AddVariableInitialization(Iterator)
            End If

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

        Public Overrides Sub EmitIL(ByVal scope As CodeGenScope)
            InitialValue.EmitIL(scope)
            Dim field = scope.Fields(Iterator.NormalizedText)
            scope.ILGenerator.Emit(OpCodes.Stsfld, field)
            Dim label As Label = scope.ILGenerator.DefineLabel()
            Dim label2 As Label = scope.ILGenerator.DefineLabel()
            Dim label3 As Label = scope.ILGenerator.DefineLabel()
            Dim label4 As Label = scope.ILGenerator.DefineLabel()
            scope.ILGenerator.MarkLabel(label2)
            scope.ILGenerator.Emit(OpCodes.Ldsfld, field)
            FinalValue.EmitIL(scope)

            If StepValue IsNot Nothing Then
                StepValue.EmitIL(scope)
                scope.ILGenerator.Emit(OpCodes.Ldc_R8, 0.0)
                scope.ILGenerator.EmitCall(OpCodes.Call, scope.TypeInfoBag.NumberToPrimitive, Nothing)
                scope.ILGenerator.EmitCall(OpCodes.Call, scope.TypeInfoBag.LessThan, Nothing)
                scope.ILGenerator.EmitCall(OpCodes.Call, scope.TypeInfoBag.PrimitiveToBoolean, Nothing)
                scope.ILGenerator.Emit(OpCodes.Brfalse, label4)
                scope.ILGenerator.EmitCall(OpCodes.Call, scope.TypeInfoBag.GreaterThanOrEqualTo, Nothing)
                scope.ILGenerator.EmitCall(OpCodes.Call, scope.TypeInfoBag.PrimitiveToBoolean, Nothing)
                scope.ILGenerator.Emit(OpCodes.Brfalse, label)
                scope.ILGenerator.Emit(OpCodes.Br, label3)
            End If

            scope.ILGenerator.MarkLabel(label4)
            scope.ILGenerator.EmitCall(OpCodes.Call, scope.TypeInfoBag.LessThanOrEqualTo, Nothing)
            scope.ILGenerator.EmitCall(OpCodes.Call, scope.TypeInfoBag.PrimitiveToBoolean, Nothing)
            scope.ILGenerator.Emit(OpCodes.Brfalse, label)
            scope.ILGenerator.MarkLabel(label3)

            For Each item In ForBody
                item.EmitIL(scope)
            Next

            scope.ILGenerator.Emit(OpCodes.Ldsfld, field)

            If StepValue IsNot Nothing Then
                StepValue.EmitIL(scope)
            Else
                scope.ILGenerator.Emit(OpCodes.Ldc_R8, 1.0)
                scope.ILGenerator.EmitCall(OpCodes.Call, scope.TypeInfoBag.NumberToPrimitive, Nothing)
            End If

            scope.ILGenerator.EmitCall(OpCodes.Call, scope.TypeInfoBag.Add, Nothing)
            scope.ILGenerator.Emit(OpCodes.Stsfld, field)
            scope.ILGenerator.Emit(OpCodes.Br, label2)
            scope.ILGenerator.MarkLabel(label)
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
