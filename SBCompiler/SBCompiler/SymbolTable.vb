Imports System.Collections.Generic
Imports System.Globalization
Imports Microsoft.SmallBasic.Expressions

Namespace Microsoft.SmallBasic
    Public Class SymbolTable

        Public ReadOnly Property Errors As List(Of [Error])

        Public ReadOnly Property InitializedVariables As New Dictionary(Of String, TokenInfo)

        Public ReadOnly Property Variables As New Dictionary(Of String, TokenInfo)

        Public ReadOnly Property Locals As New Dictionary(Of String, Expressions.IdentifierExpression)

        Public ReadOnly Property Dynamics As New Dictionary(Of String, Dictionary(Of String, TokenInfo))

        Public ReadOnly Property Subroutines As New Dictionary(Of String, TokenInfo)

        Friend Sub AddDynamic(prop As PropertyExpression)
            Dim typeNameInfo = prop.TypeName
            Dim propertyNameInfo = prop.PropertyName

            If typeNameInfo.Token = Token.Illegal OrElse propertyNameInfo.Token = Token.Illegal Then
                Return
            End If

            Dim value As TypeInfo = Nothing
            Dim typeName = typeNameInfo.NormalizedText
            Dim propertyName = propertyNameInfo.NormalizedText

            If Not _typeInfoBag.Types.TryGetValue(typeName, value) Then
                prop.IsDynamic = True
                If _Dynamics.ContainsKey(typeName) Then
                    Dim fileds = _Dynamics(typeName)
                    If Not fileds.ContainsKey(propertyName) Then
                        If prop.isSet Then fileds.Add(propertyName, propertyNameInfo)
                    End If

                ElseIf prop.isSet Then
                    _Dynamics.Add(typeName, New Dictionary(Of String, TokenInfo) From {{propertyName, propertyNameInfo}})
                    Dim subroutine = If(_Variables.ContainsKey(typeName), Nothing, Statements.SubroutineStatement.GetSubroutine(prop))
                    Dim id = New IdentifierExpression() With {
                        .Identifier = typeNameInfo,
                        .Subroutine = subroutine
                    }
                    AddVariable(id, subroutine IsNot Nothing)
                    AddVariableInitialization(typeNameInfo)
                End If
            End If
        End Sub

        Public ReadOnly Property Labels As New Dictionary(Of String, TokenInfo)

        Dim _typeInfoBag As TypeInfoBag

        Public Sub New(ByVal errors As List(Of [Error]), typeInfoBag As TypeInfoBag)
            _Errors = errors
            _typeInfoBag = typeInfoBag
            If _Errors Is Nothing Then
                _Errors = New List(Of [Error])()
            End If
        End Sub

        Public Sub CopyFrom(symbolTable As SymbolTable)
            _typeInfoBag = symbolTable._typeInfoBag
            For Each info In symbolTable._Locals
                _Locals(info.Key) = info.Value
            Next

            Copy(symbolTable._Variables, _Variables)
            Copy(symbolTable.InitializedVariables, Me.InitializedVariables)
            Copy(symbolTable.Labels, Me.Labels)
            Copy(symbolTable.Subroutines, Me.Subroutines)
        End Sub

        Private Sub Copy(fromDic As Dictionary(Of String, TokenInfo), toDic As Dictionary(Of String, TokenInfo))
            For Each info In fromDic
                toDic(info.Key) = info.Value
            Next
        End Sub

        Public Sub Reset()
            _Errors.Clear()
            _Labels.Clear()
            _Subroutines.Clear()
            _Variables.Clear()
            _Locals.Clear()
            _Dynamics.Clear()
        End Sub

        Public Sub AddVariable(variable As Expressions.IdentifierExpression, Optional isLocal As Boolean = False)
            Dim variableName = variable.Identifier.NormalizedText
            Dim Subroutine = variable.Subroutine
            Dim key = GetKey(variable)

            If _Locals.ContainsKey(key) Then Return

            If isLocal Then
                _Locals.Add(key, variable)

            ElseIf Not _Variables.ContainsKey(variableName) Then
                If Subroutine Is Nothing Then
                    _Variables.Add(variableName, variable.Identifier)
                Else
                    _Locals.Add(key, variable)
                End If
            End If
        End Sub

        Private Function GetKey(variable As IdentifierExpression) As Object
            Dim variableName = variable.Identifier.NormalizedText
            Dim Subroutine = variable.Subroutine

            If Subroutine Is Nothing Then Return variableName
            Return $"{Subroutine.Name.NormalizedText}.{variableName}"
        End Function



        Public Function IsDefined(variable As IdentifierExpression) As Boolean
            Dim var = variable.Identifier.NormalizedText
            If _Variables.ContainsKey(var) Then Return True

            Dim key = GetKey(variable)
            If _Locals.ContainsKey(key) Then
                If Not Completion.CompletionHelper.AutoCompletion Then Return True

                Dim varUse = variable.Identifier
                Dim varDeclaration = _Locals(key).Identifier
                Select Case varUse.Line
                    Case varDeclaration.Line
                        Return varUse.Column = varDeclaration.Column
                    Case Else
                        Return varUse.Line > varDeclaration.Line
                End Select
            End If

            Return False
        End Function

        Public Sub AddVariableInitialization(variable As TokenInfo)
            If Not InitializedVariables.ContainsKey(variable.NormalizedText) Then
                InitializedVariables.Add(variable.NormalizedText, variable)
            End If
        End Sub

        Public Sub AddSubroutine(subroutineName As TokenInfo, type As Token)
            Dim normalizedText = subroutineName.NormalizedText
            subroutineName.Token = type

            If _Variables.ContainsKey(normalizedText) Then
                _Variables.Remove(normalizedText)
            End If

            If Not Subroutines.ContainsKey(normalizedText) Then
                Subroutines.Add(normalizedText, subroutineName)
                Return
            End If

            Errors.Add(New [Error](subroutineName, String.Format(CultureInfo.CurrentUICulture, ResourceHelper.GetString("AnotherSubroutineExists"), New Object(0) {subroutineName.Text})))
        End Sub

        Public Sub AddLabelDefinition(ByVal label As TokenInfo)
            Dim normalizedText = label.NormalizedText

            If Not Labels.ContainsKey(normalizedText) Then
                Labels.Add(normalizedText, label)
                Return
            End If

            Errors.Add(New [Error](label, String.Format(CultureInfo.CurrentUICulture, ResourceHelper.GetString("AnotherLabelExists"), New Object(0) {label.Text})))
        End Sub
    End Class
End Namespace
