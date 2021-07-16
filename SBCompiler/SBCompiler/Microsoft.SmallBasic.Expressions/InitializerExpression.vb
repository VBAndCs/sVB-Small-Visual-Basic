Imports System
Imports System.Collections.Generic
Imports System.Globalization
Imports System.Reflection
Imports System.Reflection.Emit
Imports System.Text
Imports Microsoft.SmallBasic.Library

Namespace Microsoft.SmallBasic.Expressions
    <Serializable>
    Public Class InitializerExpression
        Inherits Expression

        Public Sub New()
            MyBase.New()
        End Sub

        Public Sub New(precedence As Integer, arguments As List(Of Expression))
            Me.Precedence = precedence
            _Arguments = arguments
        End Sub

        Public ReadOnly Property Arguments As New List(Of Expression)

        Public Overrides Sub AddSymbols(ByVal symbolTable As SymbolTable)
            For Each argument In Arguments
                argument.AddSymbols(symbolTable)
            Next
        End Sub

        Public Function Lower(leftValue As String, Optional n As Integer = 0) As String
            n += 1
            Dim tmpVar = "__tmpArray__" & n
            Dim code As New StringBuilder($"{tmpVar} = """"")
            code.AppendLine()

            For i = 0 To Arguments.Count - 1
                If TypeOf Arguments(i) Is InitializerExpression Then
                    Dim initExpr = CType(Arguments(i), InitializerExpression)
                    code.AppendLine()
                    code.AppendLine(initExpr.Lower($"{tmpVar}[{i + 1}]", n))
                Else
                    code.AppendLine($"{tmpVar}[{i + 1}] = {Arguments(i)}")
                End If
            Next

            If leftValue <> "" Then
                code.AppendLine($"{leftValue} = {tmpVar}")
            End If
            Return code.ToString()
        End Function

        Public Overrides Sub EmitIL(scope As CodeGenScope)
            Dim code = Me.Lower("")

            Dim _parser = Parser.Parse(code)
            Dim semantic As New SemanticAnalyzer(_parser, scope.TypeInfoBag)
            semantic.Analyze()
            Dim _newSymbolTable = _parser.SymbolTable
            Dim _oldSymbolTable = scope.SymbolTable
            _oldSymbolTable.CopyFrom(_newSymbolTable)

            'Build new fields
            Dim symbolTable = _parser.SymbolTable
            For Each key In symbolTable.Variables.Keys
                Dim value As FieldInfo = scope.TypeBuilder.DefineField(key, GetType(Primitive), FieldAttributes.Private Or FieldAttributes.Static)
                If Not scope.Fields.ContainsKey(key) Then scope.Fields.Add(key, value)
            Next

            ' EmitIL
            For Each item In _parser.ParseTree
                item.PrepareForEmit(scope)
            Next

            For Each item In _parser.ParseTree
                item.EmitIL(scope)
            Next

            CType(_parser.ParseTree(0), Statements.AssignmentStatement).LeftValue.EmitIL(scope)
        End Sub

        Public Overrides Function ToString() As String
            Dim stringBuilder As New StringBuilder("{")

            If Arguments.Count > 0 Then
                For Each argument In Arguments
                    stringBuilder.AppendFormat(CultureInfo.CurrentUICulture, "{0}, ", New Object(0) {argument})
                Next

                stringBuilder.Remove(stringBuilder.Length - 2, 2)
            End If

            stringBuilder.Append("}")
            Return stringBuilder.ToString()
        End Function
    End Class
End Namespace
