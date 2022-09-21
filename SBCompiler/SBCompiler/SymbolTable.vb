Imports System.Collections.Generic
Imports System.Globalization
Imports Microsoft.SmallBasic.Expressions

Namespace Microsoft.SmallBasic
    Public Class SymbolTable
        Friend AutoCompletion As Boolean
        Friend AllCommentLines As New List(Of Token)
        Public AllIdentifiers As New List(Of Token)
        Public AllStatements As New Dictionary(Of Integer, Statements.Statement)
        Public AllDynamicProperties As New Dictionary(Of String, List(Of Token))
        Public AllLibMembers As New List(Of Token)
        Public Property ControlNames As List(Of String)
        Public Property ModuleNames As Dictionary(Of String, String)

        Friend _typeInfoBag As TypeInfoBag
        Friend ReadOnly PossibleEventHandlers As New List(Of (Id As Token, index As Integer))
        Friend IsLoweredCode As Boolean

        Friend Sub AddIdentifier(identifier As Token)
            If Not IsLoweredCode Then AllIdentifiers.Add(identifier)
        End Sub

        Public ReadOnly Property Errors As List(Of [Error])

        Public ReadOnly Property InitializedVariables As New Dictionary(Of String, Token)

        Public ReadOnly Property GlobalVariables As New Dictionary(Of String, Token)

        Public ReadOnly Property LocalVariables As New Dictionary(Of String, Expressions.IdentifierExpression)

        Public ReadOnly Property Dynamics As New Dictionary(Of String, Dictionary(Of String, Token))

        Public ReadOnly Property Subroutines As New Dictionary(Of String, Token)


        Friend Sub AddDynamic(prop As PropertyExpression)
            Dim typeNameInfo = prop.TypeName
            Dim propertyNameInfo = prop.PropertyName

            If typeNameInfo.Type = TokenType.Illegal OrElse
                    propertyNameInfo.Type = TokenType.Illegal Then Return

            typeNameInfo.SymbolType = Completion.CompletionItemType.TypeName
            AddIdentifier(typeNameInfo)

            Dim value As TypeInfo = Nothing
            Dim typeName = typeNameInfo.LCaseText
            If Not prop.IsDynamic AndAlso _typeInfoBag.Types.TryGetValue(typeName, value) Then Return

            Dim propertyName = propertyNameInfo.LCaseText
            propertyNameInfo.Comment = prop.Parent.GetSummery()
            prop.PropertyName = propertyNameInfo

            If AllDynamicProperties.ContainsKey(typeName) Then
                AllDynamicProperties(typeName).Add(propertyNameInfo)
            Else
                AllDynamicProperties(typeName) = New List(Of Token) From {propertyNameInfo}
            End If

            If _Dynamics.ContainsKey(typeName) Then
                Dim fileds = _Dynamics(typeName)
                If Not fileds.ContainsKey(propertyName) AndAlso prop.isSet Then
                    fileds.Add(propertyName, propertyNameInfo)
                End If

            ElseIf prop.isSet Then
                _Dynamics.Add(
                          typeName,
                          New Dictionary(Of String, Token) From {
                                  {propertyName, propertyNameInfo}
                          }
                    )

                Dim subroutine = If(_GlobalVariables.ContainsKey(typeName), Nothing, Statements.SubroutineStatement.GetSubroutine(prop))
                Dim idExpr = New IdentifierExpression() With {
                        .Identifier = typeNameInfo,
                        .Subroutine = subroutine
                 }

                propertyNameInfo.SymbolType = Completion.CompletionItemType.DynamicPropertyName
                AddIdentifier(propertyNameInfo)

                AddVariable(idExpr, $"Dynamic {If(prop.IsDynamic, "", "(Data)")} object. Use {If(prop.IsDynamic, "`!`", ".")} to add properties to this object or to read them", subroutine IsNot Nothing)
                AddVariableInitialization(typeNameInfo)
            End If

            prop.IsDynamic = True
        End Sub

        Public Property InferedTypes As New Dictionary(Of String, VariableType)

        Public Function GetInferedType(var As Token) As VariableType
            Dim variableName = var.LCaseText
            Dim Subroutine = var.SubroutineName?.ToLower()
            If Subroutine = "" Then Return GetInferedType(variableName)

            Dim localKey = $"{Subroutine}.{variableName}"
            If _LocalVariables.ContainsKey(localKey) Then
                Return GetInferedType(localKey)
            Else
                Return GetInferedType(variableName)
            End If

        End Function

        Public Function GetInferedType(key As String) As VariableType
            If InferedTypes.ContainsKey(key) Then
                Return InferedTypes(key)
            Else
                Return VariableType.None
            End If
        End Function

        Friend Sub FixNames(typeName As Token, memberName As Token, isMethod As Boolean)
            If _typeInfoBag Is Nothing Then Return

            Dim type = GetTypeInfo(typeName)
            If type Is Nothing Then Return

            If typeName.Comment <> "" AndAlso typeName.Comment <> typeName.Text Then
                AllLibMembers.Add(typeName)
            End If

            If GetMemberInfo(memberName, type, isMethod) Is Nothing Then Return

            If memberName.Comment <> "" AndAlso memberName.Comment <> memberName.Text Then
                AllLibMembers.Add(memberName)
            End If

        End Sub

        Friend Function GetTypeInfo(ByRef typeName As Token) As TypeInfo
            If typeName.IsIllegal Then Return Nothing

            Dim type As TypeInfo
            Dim typeKey = typeName.LCaseText

            If typeKey = "me" Then
                typeName.Comment = "Me"
                type = _typeInfoBag.Types("form")

            ElseIf _typeInfoBag.Types.ContainsKey(typeKey) Then
                type = _typeInfoBag.Types(typeKey)
                typeName.Comment = type.Type.Name

            Else
                Dim varType = GetInferedType(typeName)
                Dim varTypeName = WinForms.PreCompiler.GetTypeName(varType)

                If varTypeName <> "" Then
                    typeKey = varTypeName.ToLower()
                    type = _typeInfoBag.Types(typeKey)

                ElseIf ControlNames Is Nothing Then
                    Return Nothing

                Else
                    For Each controlName In ControlNames
                        If controlName.ToLower = typeKey Then
                            typeName.Comment = controlName
                            typeKey = ModuleNames(typeKey).ToLower()
                            Exit For
                        End If
                    Next

                    If typeName.Comment = "" Then
                        Dim name = WinForms.PreCompiler.GetModuleFromVarName(typeKey)
                        If name = "" Then Return Nothing
                        type = _typeInfoBag.Types(name.ToLower())
                    Else
                        type = _typeInfoBag.Types(typeKey)
                    End If
                End If
            End If

            Return type
        End Function

        Friend Function GetMemberInfo(ByRef memberName As Token, type As TypeInfo, isMethod As Boolean) As System.Reflection.MemberInfo
            Dim memberKey = memberName.LCaseText
            Dim memberInfo As System.Reflection.MemberInfo

            If isMethod Then
                If type.Methods.ContainsKey(memberKey) Then
                    memberInfo = type.Methods(memberKey)
                    memberName.Comment = memberInfo.Name
                    Return memberInfo
                End If

            ElseIf type.Properties.ContainsKey(memberKey) Then
                memberInfo = type.Properties(memberKey)
                memberName.Comment = memberInfo.Name
                Return memberInfo

            ElseIf type.Events.ContainsKey(memberKey) Then
                memberInfo = type.Events(memberKey)
                memberName.Comment = memberInfo.Name
                Return memberInfo

            Else
                memberKey = "get" & memberKey
                If type.Methods.ContainsKey(memberKey) Then
                    memberInfo = type.Methods(memberKey)
                    memberName.Comment = memberInfo.Name.Substring(3)
                    Return memberInfo
                End If
            End If

            If type.Key <> "control" Then
                Return GetMemberInfo(memberName, _typeInfoBag.Types("control"), isMethod)
            ElseIf memberName.Comment = "" Then
                Return Nothing
            Else
                Return memberInfo
            End If

        End Function

        Public Function GetReturnValueType(typeName As Token, memberName As Token, isMethod As Boolean) As VariableType
            If _typeInfoBag Is Nothing Then Return VariableType.None

            Dim type = GetTypeInfo(typeName)
            If type Is Nothing Then Return VariableType.None

            Dim memberInfo = GetMemberInfo(memberName, type, isMethod)
            If memberInfo Is Nothing Then Return VariableType.None

            Dim attrs = memberInfo.GetCustomAttributes(GetType(WinForms.ReturnValueTypeAttribute), False)
            If attrs Is Nothing OrElse attrs.Count = 0 Then Return VariableType.None

            Return CType(attrs(0), WinForms.ReturnValueTypeAttribute).ReturnTypeValue
        End Function

        Public ReadOnly Labels As New Dictionary(Of String, Token)

        Public Sub New(errors As List(Of [Error]), typeInfoBag As TypeInfoBag)
            _Errors = errors
            _typeInfoBag = typeInfoBag
            If _Errors Is Nothing Then
                _Errors = New List(Of [Error])()
            End If
        End Sub

        Public Sub CopyFrom(symbolTable As SymbolTable)
            _typeInfoBag = symbolTable._typeInfoBag
            For Each info In symbolTable._LocalVariables
                _LocalVariables(info.Key) = info.Value
            Next

            Copy(symbolTable._GlobalVariables, _GlobalVariables)
            Copy(symbolTable.InitializedVariables, Me.InitializedVariables)
            Copy(symbolTable.Labels, Me.Labels)
            Copy(symbolTable.Subroutines, Me.Subroutines)
        End Sub

        Private Sub Copy(fromDic As Dictionary(Of String, Token), toDic As Dictionary(Of String, Token))
            For Each info In fromDic
                toDic(info.Key) = info.Value
            Next
        End Sub

        Public Sub Reset()
            _Errors.Clear()
            Labels.Clear()
            _Subroutines.Clear()
            _GlobalVariables.Clear()
            _LocalVariables.Clear()
            _Dynamics.Clear()
            InferedTypes.Clear()

            AllCommentLines.Clear()
            AllStatements.Clear()
            AllIdentifiers.Clear()
            AllDynamicProperties.Clear()
            AllLibMembers.Clear()
            PossibleEventHandlers.Clear()
        End Sub

        Public Function AddVariable(
                         variable As Expressions.IdentifierExpression,
                         comment As String,
                         Optional isLocal As Boolean = False
                   ) As String

            Dim variableName = variable.Identifier.LCaseText

            If variableName = "_" Then
                Errors.Add(New [Error](variable.Identifier, "_ is not a valid name"))
                Return ""
            End If

            Dim Subroutine = variable.Subroutine
            If comment <> "" Then variable.Identifier.Comment = comment

            Dim newGlobal = Not _GlobalVariables.ContainsKey(variableName)

            If isLocal Then   ' There can be a local var and a global var with the same name. 
                Return AddLocalVar(variable)

            ElseIf Subroutine Is Nothing Then
                If newGlobal Then
                    _GlobalVariables.Add(variableName, variable.Identifier)
                    Return variableName
                End If

            ElseIf newGlobal Then
                Return AddLocalVar(variable)
            End If

            Return ""
        End Function

        Private Function AddLocalVar(variable As IdentifierExpression) As String
            Dim key = GetKey(variable)
            If _LocalVariables.ContainsKey(key) Then Return ""
            _LocalVariables.Add(key, variable)
            Return key
        End Function

        Public Function IsGlobalVar(variable As IdentifierExpression) As Boolean
            If IsLocalVar(variable) Then Return False
            Return _GlobalVariables.ContainsKey(variable.Identifier.LCaseText)
        End Function

        Public Function IsGlobalVar(variable As Token) As Boolean
            If IsLocalVar(variable) Then Return False
            Return _GlobalVariables.ContainsKey(variable.LCaseText)
        End Function

        Public Function IsLocalVar(variable As IdentifierExpression) As Boolean
            Return _LocalVariables.ContainsKey(GetKey(variable))
        End Function

        Public Function IsLocalVar(variable As Token) As Boolean
            Return _LocalVariables.ContainsKey(GetKey(variable))
        End Function

        Private Function GetKey(variable As IdentifierExpression) As String
            Dim variableName = variable.Identifier.LCaseText
            Dim Subroutine = variable.Subroutine

            If Subroutine Is Nothing Then Return variableName
            Return $"{Subroutine.Name.LCaseText}.{variableName}"
        End Function

        Private Function GetKey(identifier As Token) As String
            Dim variableName = identifier.LCaseText
            Dim Subroutine = identifier.SubroutineName?.ToLower()
            If Subroutine = "" Then Return variableName
            Return $"{Subroutine}.{variableName}"
        End Function

        Public Function IsDefined(variable As IdentifierExpression) As Boolean
            Return IsDefined(variable.Identifier)
        End Function


        Public Function IsDefined(identifier As Token, Optional usedAfterDefind As Boolean = False) As Boolean
            Dim var = identifier.LCaseText
            Dim varDeclaration As Token
            Dim key = GetKey(identifier)

            If _LocalVariables.ContainsKey(key) Then
                varDeclaration = _LocalVariables(key).Identifier
            ElseIf _GlobalVariables.ContainsKey(var) Then
                varDeclaration = _GlobalVariables(var)
                If identifier.SubroutineName <> "" Then Return True
            Else
                Return usedAfterDefind
            End If

            If Not usedAfterDefind AndAlso Not AutoCompletion Then Return True

            Select Case identifier.Line
                Case varDeclaration.Line
                    Return identifier.Column = varDeclaration.Column
                Case Else
                    Return identifier.Line > varDeclaration.Line
            End Select

        End Function

        Public Function UsedBeforeDefind(identifier As Token)
            Select Case identifier.SymbolType
                Case Completion.CompletionItemType.LocalVariable, Completion.CompletionItemType.GlobalVariable
                    Return Not IsDefined(identifier, True)
            End Select
            Return False
        End Function

        Public Sub AddVariableInitialization(variable As Token)
            If Not InitializedVariables.ContainsKey(variable.LCaseText) Then
                InitializedVariables.Add(variable.LCaseText, variable)
            End If
        End Sub

        Public Sub AddSubroutine(subroutine As Token, type As TokenType)
            Dim name = subroutine.LCaseText
            subroutine.Type = type

            If _GlobalVariables.ContainsKey(name) Then
                _GlobalVariables.Remove(name)
            End If

            If Subroutines.ContainsKey(name) Then
                Errors.Add(New [Error](subroutine, String.Format(CultureInfo.CurrentUICulture, ResourceHelper.GetString("AnotherSubroutineExists"), New Object(0) {subroutine.Text})))
            Else
                Subroutines.Add(name, subroutine)
            End If

        End Sub

        Public Sub AddLabelDefinition(label As Token)
            Dim labelName = label.LCaseText
            If labelName = "_" Then
                Errors.Add(New [Error](label, "_ is not a valid name"))
                Return
            End If

            If Labels.ContainsKey(labelName) Then
                Errors.Add(New [Error](label, String.Format(CultureInfo.CurrentUICulture, ResourceHelper.GetString("AnotherLabelExists"), New Object(0) {label.Text})))
            Else
                Labels.Add(labelName, label)
            End If

        End Sub


    End Class


End Namespace
