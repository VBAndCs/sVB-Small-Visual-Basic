Imports System.Reflection
Imports System.Reflection.Emit
Imports Microsoft.SmallVisualBasic.Completion
Imports Microsoft.SmallVisualBasic.Library
Imports Microsoft.SmallVisualBasic.Statements

Namespace Microsoft.SmallVisualBasic.Expressions
    <Serializable>
    Public Class PropertyExpression
        Inherits Expression

        Public Property TypeName As Token
        Public Property PropertyName As Token

        Public IsDynamic As Boolean
        Public isSet As Boolean
        Public IsEvent As Boolean

        Public ReadOnly Property DynamicKey As String
            Get
                Dim dynTypeName = CompletionHelper.TrimData(TypeName.Text)
                Return $"dynprop.{dynTypeName}.{PropertyName.LCaseText}"
            End Get
        End Property


        Public Overrides Sub AddSymbols(symbolTable As SymbolTable)
            MyBase.AddSymbols(symbolTable)
            _TypeName.Parent = Me.Parent
            _PropertyName.Parent = Me.Parent

            Dim name = TypeName.LCaseText
            _TypeName.SymbolType = Completion.CompletionItemType.GlobalVariable
            symbolTable.AddIdentifier(_TypeName)

            If IsDynamic OrElse name.StartsWith("data") Or name.EndsWith("data") Then
                symbolTable.AddDynamic(Me)
            End If

            ' IsDynamic can change after calling AddDynamic
            If IsDynamic Then
                _PropertyName.SymbolType = CompletionItemType.DynamicPropertyName
            Else
                _PropertyName.SymbolType = CompletionItemType.PropertyName
                symbolTable.FixNames(_TypeName, _PropertyName, False)
            End If

            symbolTable.AddIdentifier(_PropertyName)
        End Sub

        Private Shared dynCounter As Integer = 0

        Public Overrides Sub EmitIL(scope As CodeGenScope)
            If scope.ForGlobalHelp Then
                LiteralExpression.Zero.EmitIL(scope)
                Return
            End If

            If IsDynamic Then
                dynCounter += 1
                Dim code = $"_sVB_dynamic_Data_{dynCounter}={TypeName.Text}[""{PropertyName.Text}""]"
                Dim subroutine = SubroutineStatement.GetSubroutine(Me)
                If subroutine Is Nothing Then subroutine = SubroutineStatement.Current
                Parser.ParseAndEmit(code, subroutine, scope, StartToken.Line).EmitIL(scope)
            Else
                Dim typeInfo = scope.TypeInfoBag.Types(TypeName.LCaseText)
                Dim propertyInfo = typeInfo.Properties(PropertyName.LCaseText)
                Dim getMethod = propertyInfo.GetGetMethod()
                scope.ILGenerator.EmitCall(OpCodes.Call, getMethod, Nothing)
            End If
        End Sub

        Public Overrides Function ToString() As String
            Return $"{TypeName.Text}{If(IsDynamic, "!", ".")}{PropertyName.Text}"
        End Function

        Public Overrides Function InferType(symbolTable As SymbolTable) As VariableType
            If IsDynamic Then
                Return symbolTable.GetInferedType(DynamicKey)
            Else
                Return symbolTable.GetReturnValueType(_TypeName, _PropertyName, False)
            End If
        End Function

        Public Overrides Function Evaluate(runner As Engine.ProgramRunner) As Primitive
            Dim tName = _TypeName.LCaseText
            If IsDynamic Then
                Dim arrExpr As New ArrayExpression() With {
                      .LeftHand = New IdentifierExpression() With {.Identifier = _TypeName},
                      .Indexer = New LiteralExpression($"""{_PropertyName.Text}"""),
                      .Parent = Me.Parent
                }
                Return arrExpr.Evaluate(runner)

            ElseIf tName = "global" Then
                Dim result = runner.GetGlobalField(_PropertyName.LCaseText)
                Return If(result, New Primitive("???"))

            ElseIf runner.TypeInfoBag.Types.ContainsKey(tName) Then
                Dim typeInfo = runner.TypeInfoBag.Types(tName)
                Dim propKey = _PropertyName.LCaseText
                If Not typeInfo.Properties.ContainsKey(propKey) Then Return New Primitive("???")

                Dim propertyInfo = typeInfo.Properties(propKey)
                Return CType(propertyInfo.GetValue(Nothing, Nothing), Primitive)

            Else
                Dim type = runner.SymbolTable.GetTypeInfo(_TypeName)
                Dim memberInfo = runner.SymbolTable.GetMemberInfo(_PropertyName, type, False)
                If memberInfo Is Nothing Then Return New Primitive("???")

                Dim propInfo = TryCast(memberInfo, PropertyInfo)
                If propInfo IsNot Nothing Then
                    Return CType(propInfo.GetValue(Nothing, Nothing), Primitive)
                End If

                Dim methodInfo = TryCast(memberInfo, MethodInfo)
                If methodInfo IsNot Nothing Then
                    Dim key = runner.GetKey(_TypeName)
                    If Not runner.Fields.ContainsKey(key) Then Return New Primitive("This object is not set yet")
                    Return CType(methodInfo.Invoke(Nothing, New Object() {runner.Fields(key)}), Primitive)
                End If

                Return New Primitive("Event Handler")
            End If
        End Function

        Public Overrides Function ToVB() As String
            If IsDynamic Then
                Return $"{TypeName.Text}(""{PropertyName.Text}"")"
            End If
            Return $"{TypeName.Text}.{PropertyName.Text}"
        End Function
    End Class
End Namespace
