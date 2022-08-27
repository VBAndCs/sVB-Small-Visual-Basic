Imports System.Collections.Generic
Imports System.Globalization
Imports Microsoft.SmallBasic.Expressions

Namespace Microsoft.SmallBasic
    Public Class SymbolTable
        Friend autoCompletion As Boolean
        Friend AllCommentLines As New List(Of Token)
        Public AllIdentifiers As New List(Of Token)
        Public AllStatements As New Dictionary(Of Integer, Statements.Statement)
        Public AllDynamicProperties As New Dictionary(Of String, List(Of Token))
        Public AllLibMembers As New List(Of Token)
        Public ControlNames As List(Of String)
        Public ModuleNames As Dictionary(Of String, String)

        Dim _typeInfoBag As TypeInfoBag
        Friend ReadOnly PossibleEventHandlers As New List(Of (Id As Token, index As Integer))
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
            AllIdentifiers.Add(typeNameInfo)

            Dim value As TypeInfo = Nothing
            Dim typeName = typeNameInfo.NormalizedText
            If Not prop.IsDynamic AndAlso _typeInfoBag.Types.TryGetValue(typeName, value) Then Return

            Dim propertyName = propertyNameInfo.NormalizedText
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
                AllIdentifiers.Add(propertyNameInfo)

                AddVariable(idExpr, $"Dynamic {If(prop.IsDynamic, "", "(Data)")} object. Use {If(prop.IsDynamic, "`!`", ".")} to add properties to this object or to read them", subroutine IsNot Nothing)
                AddVariableInitialization(typeNameInfo)
            End If

            prop.IsDynamic = True
        End Sub

        Friend Sub FixNames(typeName As Token, memberName As Token, isMethod As Boolean)
            If ControlNames Is Nothing Then Return
            If typeName.IsIllegal Then Return

            Dim type As TypeInfo
            Dim typeKey = typeName.NormalizedText

            If typeKey = "me" Then
                typeName.Comment = "Me"
                type = _typeInfoBag.Types("form")

            ElseIf _typeInfoBag.Types.ContainsKey(typeKey) Then
                type = _typeInfoBag.Types(typeKey)
                typeName.Comment = type.Type.Name

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
                    If name = "" Then Return
                    type = _typeInfoBag.Types(name.ToLower())
                Else
                    type = _typeInfoBag.Types(typeKey)
                End If

            End If

            If typeName.Comment <> "" AndAlso typeName.Comment <> typeName.Text Then
                AllLibMembers.Add(typeName)
            End If

            Dim firstTime = True

CouldBeAControl:
            Dim memberKey = memberName.NormalizedText
            If isMethod Then
                If type.Methods.ContainsKey(memberKey) Then
                    memberName.Comment = type.Methods(memberKey).Name
                    firstTime = False
                End If

            ElseIf type.Properties.ContainsKey(memberKey) Then
                memberName.Comment = type.Properties(memberKey).Name
                firstTime = False

            ElseIf type.Events.ContainsKey(memberKey) Then
                memberName.Comment = type.Events(memberKey).Name
                firstTime = False

            Else
                memberKey = "get" & memberKey
                If type.Methods.ContainsKey(memberKey) Then
                    memberName.Comment = type.Methods(memberKey).Name.Substring(3)
                    firstTime = False
                End If
            End If

            If firstTime Then
                type = _typeInfoBag.Types("control")
                firstTime = False
                GoTo CouldBeAControl

            ElseIf typeName.Comment = "" Then
                Return
            End If

            If memberName.Comment <> memberName.Text Then
                AllLibMembers.Add(memberName)
            End If

        End Sub

        Public ReadOnly Property Labels As New Dictionary(Of String, Token)

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
            _Labels.Clear()
            _Subroutines.Clear()
            _GlobalVariables.Clear()
            _LocalVariables.Clear()
            _Dynamics.Clear()

            AllCommentLines.Clear()
            AllStatements.Clear()
            AllIdentifiers.Clear()
            AllDynamicProperties.Clear()
            AllLibMembers.Clear()
            PossibleEventHandlers.Clear()
        End Sub

        Public Sub AddVariable(
                         variable As Expressions.IdentifierExpression,
                         comment As String,
                         Optional isLocal As Boolean = False
                   )

            Dim variableName = variable.Identifier.NormalizedText

            If variableName = "_" Then
                Errors.Add(New [Error](variable.Identifier, "_ is not a valid name"))
                Return
            End If

            Dim Subroutine = variable.Subroutine
            If comment <> "" Then variable.Identifier.Comment = comment

            Dim newGlobal = Not _GlobalVariables.ContainsKey(variableName)

            If isLocal Then   ' There can be a local var and a global var with the same name. 
                AddLocalVar(variable)

            ElseIf Subroutine Is Nothing Then
                If newGlobal Then _GlobalVariables.Add(variableName, variable.Identifier)

            ElseIf newGlobal Then
                AddLocalVar(variable)
            End If
        End Sub

        Private Sub AddLocalVar(variable As IdentifierExpression)
            Dim key = GetKey(variable)
            If _LocalVariables.ContainsKey(key) Then Return
            _LocalVariables.Add(key, variable)
        End Sub

        Public Function IsGlobalVar(variable As IdentifierExpression) As Boolean
            If IsLocalVar(variable) Then Return False
            Return _GlobalVariables.ContainsKey(variable.Identifier.NormalizedText)
        End Function

        Public Function IsGlobalVar(variable As Token) As Boolean
            If IsLocalVar(variable) Then Return False
            Return _GlobalVariables.ContainsKey(variable.NormalizedText)
        End Function

        Public Function IsLocalVar(variable As IdentifierExpression) As Boolean
            Return _LocalVariables.ContainsKey(GetKey(variable))
        End Function

        Public Function IsLocalVar(variable As Token) As Boolean
            Return _LocalVariables.ContainsKey(GetKey(variable))
        End Function

        Private Function GetKey(variable As IdentifierExpression) As String
            Dim variableName = variable.Identifier.NormalizedText
            Dim Subroutine = variable.Subroutine

            If Subroutine Is Nothing Then Return variableName
            Return $"{Subroutine.Name.NormalizedText}.{variableName}"
        End Function

        Private Function GetKey(identifier As Token) As String
            Dim variableName = identifier.NormalizedText
            Dim Subroutine = identifier.SubroutineName?.ToLower()
            If Subroutine = "" Then Return variableName
            Return $"{Subroutine}.{variableName}"
        End Function

        Public Function IsDefined(variable As IdentifierExpression) As Boolean
            Dim var = variable.Identifier.NormalizedText
            If _GlobalVariables.ContainsKey(var) Then Return True

            Dim key = GetKey(variable)
            If _LocalVariables.ContainsKey(key) Then
                If Not autoCompletion Then Return True

                Dim varUse = variable.Identifier
                Dim varDeclaration = _LocalVariables(key).Identifier
                Select Case varUse.Line
                    Case varDeclaration.Line
                        Return varUse.Column = varDeclaration.Column
                    Case Else
                        Return varUse.Line > varDeclaration.Line
                End Select
            End If

            Return False
        End Function

        Public Sub AddVariableInitialization(variable As Token)
            If Not InitializedVariables.ContainsKey(variable.NormalizedText) Then
                InitializedVariables.Add(variable.NormalizedText, variable)
            End If
        End Sub

        Public Sub AddSubroutine(subroutine As Token, type As TokenType)
            Dim name = subroutine.NormalizedText
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
            Dim labelName = label.NormalizedText
            If labelName = "_" Then
                Errors.Add(New [Error](label, "_ is not a valid name"))
                Return
            End If

            If Not Labels.ContainsKey(labelName) Then
                Labels.Add(labelName, label)
            Else
                Errors.Add(New [Error](label, String.Format(CultureInfo.CurrentUICulture, ResourceHelper.GetString("AnotherLabelExists"), New Object(0) {label.Text})))
            End If

        End Sub
    End Class
End Namespace
