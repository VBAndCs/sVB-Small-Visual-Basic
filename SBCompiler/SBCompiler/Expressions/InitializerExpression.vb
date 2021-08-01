Imports System
Imports System.Collections.Generic
Imports System.Globalization
Imports System.Reflection
Imports System.Reflection.Emit
Imports System.Text
Imports Microsoft.SmallBasic.Library
Imports Microsoft.SmallBasic.Statements

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
            MyBase.AddSymbols(symbolTable)

            For Each argument In Arguments
                argument.Parent = Me.Parent
                argument.AddSymbols(symbolTable)
            Next
        End Sub

        Private Function Lower(leftValue As String, Optional n As Integer = 0) As String
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

        Public Function LowerAndEmit(leftValue As String, scope As CodeGenScope) As Expression
            Dim code = Me.Lower(leftValue)
            Statements.SubroutineStatement.Current = GetSubroutine(Me)
            Dim _parser = Parser.Parse(code, scope.SymbolTable)
            Statements.SubroutineStatement.Current = Nothing

            Dim semantic As New SemanticAnalyzer(_parser, scope.TypeInfoBag)
            semantic.Analyze()

            'Build new fields
            For Each key In _parser.SymbolTable.Variables.Keys
                If Not scope.Fields.ContainsKey(key) Then
                    Dim value As FieldInfo = scope.TypeBuilder.DefineField(key, GetType(Primitive), FieldAttributes.Private Or FieldAttributes.Static)
                    scope.Fields.Add(key, value)
                End If
            Next

            ' EmitIL
            For Each item In _parser.ParseTree
                item.PrepareForEmit(scope)
            Next

            For Each item In _parser.ParseTree
                item.EmitIL(scope)
            Next

            Return CType(_parser.ParseTree(0), Statements.AssignmentStatement).LeftValue
        End Function

        Public Overrides Sub EmitIL(scope As CodeGenScope)
            LowerAndEmit("", scope).EmitIL(scope)
        End Sub



        Private Function GetSubroutine(expression As Expression) As SubroutineStatement
            Dim parent = expression.Parent
            Do Until parent Is Nothing OrElse TypeOf parent Is SubroutineStatement
                parent = parent.Parent
            Loop
            Return parent
        End Function

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
