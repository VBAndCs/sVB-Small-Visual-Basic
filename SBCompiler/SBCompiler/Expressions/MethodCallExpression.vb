Imports System.Globalization
Imports System.Reflection
Imports System.Reflection.Emit
Imports System.Text
Imports Microsoft.SmallVisualBasic.Library

Namespace Microsoft.SmallVisualBasic.Expressions
    <Serializable>
    Public Class MethodCallExpression
        Inherits Expression

        Public Sub New(startToken As Token, precedence As Integer, typeName As Token, methodName As Token, arguments As List(Of Expression))
            Me.StartToken = startToken
            Me.Precedence = precedence
            _TypeName = typeName
            _MethodName = methodName
            _Arguments = arguments
        End Sub

        Public Property TypeName As Token
        Public Property MethodName As Token
        Public Property OuterSubroutine As Statements.SubroutineStatement

        Public ReadOnly Property Arguments As New List(Of Expression)

        Public Overrides Sub AddSymbols(symbolTable As SymbolTable)
            MyBase.AddSymbols(symbolTable)
            _TypeName.Parent = Me.Parent
            _MethodName.Parent = Me.Parent

            For Each argument In Arguments
                argument.Parent = Me.Parent
                argument.AddSymbols(symbolTable)
            Next

            _TypeName.SymbolType = Completion.CompletionItemType.GlobalVariable
            symbolTable.AddIdentifier(_TypeName)

            _MethodName.SymbolType = If(_TypeName.IsIllegal, Completion.CompletionItemType.SubroutineName, Completion.CompletionItemType.MethodName)
            symbolTable.AddIdentifier(_MethodName)

            symbolTable.FixNames(_TypeName, _MethodName, True)
        End Sub

        Public Overrides Sub EmitIL(scope As CodeGenScope)
            If scope.ForGlobalHelp Then
                LiteralExpression.Zero.EmitIL(scope)
                Return
            End If

            If TypeName.Type = TokenType.Illegal Then ' Function Call
                Dim subroutine As New Statements.SubroutineCallStatement() With {
                    .Name = _MethodName,
                    .Args = _Arguments,
                    .OuterSubroutine = _OuterSubroutine,
                    .KeepReturnValue = True
                }
                subroutine.EmitIL(scope)

            Else
                Dim methodInfo = GetMethodInfo(scope)
                For Each argument In Arguments
                    argument.EmitIL(scope)
                Next
                scope.ILGenerator.EmitCall(OpCodes.Call, methodInfo, Nothing)
            End If
        End Sub

        Public Function GetMethodInfo(scope As CodeGenScope) As MethodInfo
            Dim typeInfo = scope.TypeInfoBag.Types(_TypeName.LCaseText)
            Return typeInfo.Methods(_MethodName.LCaseText)
        End Function

        Public Overrides Function ToString() As String
            Dim sb As New StringBuilder()

            If Not _TypeName.IsIllegal Then
                sb.Append($"{_TypeName.Text}.")
            End If

            sb.Append(_MethodName.Text & "(")

            If _Arguments.Count > 0 Then
                For Each arg In _Arguments
                    sb.Append(arg.ToString())
                    sb.Append(", ")
                Next

                sb.Remove(sb.Length - 2, 2)
            End If

            sb.Append(")")
            Return sb.ToString()
        End Function

        Public Overrides Function InferType(symbolTable As SymbolTable) As VariableType
            If _TypeName.IsIllegal Then
                Dim funcName = _MethodName.LCaseText
                Return symbolTable.GetInferedType(funcName)
            Else
                Return symbolTable.GetReturnValueType(_TypeName, _MethodName, True)
            End If

        End Function

        Public Overrides Function Evaluate(runner As Engine.ProgramRunner) As Primitive
            Dim tName = _TypeName.LCaseText
            Dim mName = _MethodName.LCaseText
            Dim isGlobalFunc = (tName = "global")

            If TypeName.IsIllegal OrElse isGlobalFunc Then ' Function Call
                Dim subroutine As New Statements.SubroutineCallStatement() With {
                    .Name = _MethodName,
                    .Args = _Arguments,
                    .IsGlobalFunc = isGlobalFunc
                }

                subroutine.Execute(runner)

                Dim retKey = $"{MName}.return"
                If isGlobalFunc Then
                    Return runner.GetGlobalField(retKey)
                Else
                    Return runner.Fields(retKey)
                End If
            End If

            Dim args As New List(Of Object)()
            For Each argument In _Arguments
                args.Add(argument.Evaluate(runner))
            Next

            If tName = "forms" Then
                If mName = "showform" Then
                    Dim formName = CType(args(0), Library.Primitive)
                    Dim argsArr = CType(args(1), Library.Primitive)
                    WinForms.Forms.DoShowForm(
                        formName, argsArr, Sub() runner.RunForm(formName))
                    Return formName

                ElseIf mName = "showdialog" Then
                    Dim formName = CType(args(0), Library.Primitive)
                    Dim argsArr = CType(args(1), Library.Primitive)
                    Return WinForms.Forms.DoShowDialog(
                        formName, argsArr, Sub() runner.RunForm(formName))
                End If

            ElseIf tName = "form" AndAlso mName = "showchildform" Then
                Dim parentFormName = CType(args(0), Library.Primitive)
                Dim childFormName = CType(args(1), Library.Primitive)
                Dim argsArr = CType(args(2), Library.Primitive)

                WinForms.Form.DoShowChildForm(
                    parentFormName,
                    childFormName,
                    argsArr,
                    Sub() runner.RunForm(childFormName)
                )
                Return childFormName
            End If

            Dim typeInfo = runner.TypeInfoBag.Types(tName)
            Dim methodInfo = typeInfo.Methods(_MethodName.LCaseText)
            Return CType(methodInfo.Invoke(Nothing, args.ToArray()), Primitive)
        End Function

    End Class
End Namespace
