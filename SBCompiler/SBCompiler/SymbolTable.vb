Imports System.Globalization
Imports Microsoft.SmallBasic
Imports Microsoft.SmallVisualBasic.Expressions
Imports TokenDictionary = System.Collections.Generic.Dictionary(Of String, Microsoft.SmallVisualBasic.Token)

Namespace Microsoft.SmallVisualBasic
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
        Friend Property IsLoweredCode As Boolean

        Friend Sub AddIdentifier(identifier As Token)
            If Not IsLoweredCode AndAlso Not identifier.Hidden Then
                AllIdentifiers.Add(identifier)
            End If
        End Sub

        Public ReadOnly Property Errors As List(Of [Error])

        Public ReadOnly Property InitializedVariables As New TokenDictionary

        Public ReadOnly Property GlobalVariables As New TokenDictionary

        Public ReadOnly Property LocalVariables As New Dictionary(Of String, IdentifierExpression)

        Public ReadOnly Property Dynamics As New Dictionary(Of String, TokenDictionary)

        Public ReadOnly Property Subroutines As New TokenDictionary


        Friend Sub AddDynamic(prop As PropertyExpression)
            Dim typeNameInfo = prop.TypeName
            Dim propNameInfo = prop.PropertyName

            If typeNameInfo.Type = TokenType.Illegal OrElse
                    propNameInfo.Type = TokenType.Illegal Then Return

            typeNameInfo.SymbolType = Completion.CompletionItemType.TypeName
            AddIdentifier(typeNameInfo)

            Dim typeName = typeNameInfo.LCaseText
            If Not prop.IsDynamic AndAlso _typeInfoBag.Types.ContainsKey(typeName) Then
                Return
            End If

            Dim propName = propNameInfo.LCaseText
            propNameInfo.Comment = prop.Parent.GetSummery()
            prop.PropertyName = propNameInfo

            If AllDynamicProperties.ContainsKey(typeName) Then
                AllDynamicProperties(typeName).Add(propNameInfo)
            Else
                AllDynamicProperties(typeName) = New List(Of Token) From {propNameInfo}
            End If

            If propNameInfo.Comment = "" Then
                propNameInfo.Comment = GetParentComment(typeName, propName)
            End If

            If _Dynamics.ContainsKey(typeName) Then
                Dim fileds = _Dynamics(typeName)
                If Not fileds.ContainsKey(propName) AndAlso prop.isSet Then
                    fileds.Add(propName, propNameInfo)
                End If

            ElseIf prop.isSet Then
                _Dynamics.Add(
                    typeName,
                    New TokenDictionary From {
                         {propName, propNameInfo}
                    }
                )
            End If

            If prop.isSet Then
                Dim subroutine = If(_GlobalVariables.ContainsKey(typeName), Nothing, Statements.SubroutineStatement.GetSubroutine(prop))
                Dim idExpr = New IdentifierExpression() With {
                        .Identifier = typeNameInfo,
                        .Subroutine = subroutine
                }

                propNameInfo.SymbolType = Completion.CompletionItemType.DynamicPropertyName
                AddIdentifier(propNameInfo)

                AddVariable(
                    idExpr,
                    $"`{typeNameInfo.Text}` is an array that works as a {If(prop.IsDynamic, "dynamic", "data")} object. Use {If(prop.IsDynamic, "`!`", ".")} to add properties to this object or to read them.",
                    subroutine IsNot Nothing
                )
                AddVariableInitialization(typeNameInfo)
            End If

            prop.IsDynamic = True
        End Sub

        Function GetParentComment(
                           typeName As String,
                           propName As String
                       )

            Dim propComment = ""
            For Each dynType In _Dynamics
                Dim key = dynType.Key
                If key <> typeName AndAlso (
                                typeName.StartsWith(key) OrElse typeName.EndsWith(key)
                            ) Then
                    For Each dynProp In dynType.Value
                        If dynProp.Key = propName Then
                            Dim comment = dynProp.Value.Comment
                            If comment <> "" Then
                                propComment = comment
                                Exit For
                            End If
                        End If
                    Next
                End If
            Next

            Return propComment
        End Function

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
                Return VariableType.Any
            End If
        End Function

        Friend Sub FixNames(
                        typeName As Token,
                        memberName As Token,
                        isMethod As Boolean
                    )

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
                typeName.Comment = type.Name

            ElseIf ControlNames Is Nothing Then
                Return GetTypeInfoFromInfered(typeName)

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
                    If name = "" Then Return GetTypeInfoFromInfered(typeName)
                    type = _typeInfoBag.Types(name.ToLower())
                Else
                    type = _typeInfoBag.Types(typeKey)
                End If
            End If

            Return type
        End Function

        Private Function GetTypeInfoFromInfered(typeName As Token) As TypeInfo
            Dim varType = GetInferedType(typeName)
            Dim varTypeName = WinForms.PreCompiler.GetTypeName(varType)

            If varTypeName <> "" Then
                Dim typeKey = varTypeName.ToLower()
                Return _typeInfoBag.Types(typeKey)
            End If

            Return Nothing
        End Function

        Friend Function GetMemberInfo(
                          ByRef memberName As Token,
                          type As TypeInfo,
                          isMethod As Boolean
                    ) As System.Reflection.MemberInfo

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
            If _typeInfoBag Is Nothing Then Return VariableType.Any

            Dim type = GetTypeInfo(typeName)
            If type Is Nothing Then Return VariableType.Any

            Dim memberInfo = GetMemberInfo(memberName, type, isMethod)
            If memberInfo Is Nothing Then Return VariableType.Any

            Dim attrs = memberInfo.GetCustomAttributes(GetType(WinForms.ReturnValueTypeAttribute), False)
            If attrs Is Nothing OrElse attrs.Count = 0 Then Return VariableType.Any

            Return CType(attrs(0), WinForms.ReturnValueTypeAttribute).ReturnTypeValue
        End Function

        Public ReadOnly Labels As New TokenDictionary

        Public Sub New(errors As List(Of [Error]))
            _Errors = errors
            _typeInfoBag = Compiler.TypeInfoBag
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

        Private Sub Copy(fromDic As TokenDictionary, toDic As TokenDictionary)
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
            Dim newGlobal = Not _GlobalVariables.ContainsKey(variableName)

            ' if var is invalid, we will still add it to the dictionary not to break intellisense
            ValidateVariableName(variable.Identifier)

            Dim Subroutine = variable.Subroutine
            If comment <> "" Then variable.Identifier.Comment = comment

            If isLocal Then   ' There can be a local var and a global var with the same name. 
                Return AddLocalVar(variable)

            ElseIf Subroutine Is Nothing Then
                If newGlobal Then
                    _GlobalVariables.Add(variableName, variable.Identifier)
                    Return variableName
                End If

            ElseIf newGlobal Then
                If _ControlNames IsNot Nothing Then
                    ' In design time, generated code is not added to the code file,
                    ' hence control names are not declared as variabls yet!
                    ' so, we need to add them here to make intellisense work correctly.
                    For Each controlName In _ControlNames
                        If variableName = controlName.ToLower Then
                            _GlobalVariables.Add(variableName, variable.Identifier)
                            Return variableName
                        End If
                    Next
                End If

                Return AddLocalVar(variable)
                End If

                Return ""
        End Function

        Private Sub ValidateVariableName(name As Token)
            Dim msg = $"{name.Text} is not a valid identifier name"
            Dim variableName = name.LCaseText
            Select Case variableName
                Case "_"
                    Errors.Add(New [Error](name, msg))
                Case "global"
                    Errors.Add(New [Error](name, msg & " because it is a keyword."))
                Case Else
                    If _typeInfoBag.Types.ContainsKey(variableName) Then
                        Dim type = _typeInfoBag.Types(variableName)
                        If type.Type.Assembly.FullName.StartsWith("SmallVisualBasicLibrary,") Then
                            Errors.Add(New [Error](name, msg & " because it is a sVB type."))
                        End If
                    End If
            End Select
        End Sub

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

        Friend Function GetKey(identifier As Token) As String
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
            ValidateVariableName(subroutine)

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
            ValidateVariableName(label)

            Dim labelName = label.LCaseText
            If Labels.ContainsKey(labelName) Then
                Errors.Add(New [Error](label, String.Format(CultureInfo.CurrentUICulture, ResourceHelper.GetString("AnotherLabelExists"), New Object(0) {label.Text})))
            Else
                Labels.Add(labelName, label)
            End If

        End Sub

    End Class


End Namespace
