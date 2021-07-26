Imports System.Collections.Generic
Imports System.Globalization

Namespace Microsoft.SmallBasic
    Public Class SymbolTable

        Public ReadOnly Property Errors As List(Of [Error])

        Public ReadOnly Property InitializedVariables As New Dictionary(Of String, TokenInfo)

        Public ReadOnly Property Variables As New Dictionary(Of String, TokenInfo)

        Public ReadOnly Property Locals As New Dictionary(Of String, Expressions.IdentifierExpression)

        Public ReadOnly Property Subroutines As New Dictionary(Of String, TokenInfo)

        Public ReadOnly Property Labels As New Dictionary(Of String, TokenInfo)

        Public Sub New(ByVal errors As List(Of [Error]))
            _errors = errors

            If _errors Is Nothing Then
                _errors = New List(Of [Error])()
            End If
        End Sub

        Public Sub CopyFrom(symbolTable As SymbolTable)
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
            _errors.Clear()
            _labels.Clear()
            _subroutines.Clear()
            _Variables.Clear()
            _locals.clear()
        End Sub

        Public Sub AddVariable(variable As Expressions.IdentifierExpression, Optional isLocal As Boolean = False)
            Dim Subroutine = variable.Subroutine
            Dim variableName = variable.Identifier.NormalizedText
            Dim key = ""

            If Subroutine Is Nothing Then
                key = variableName
            Else
                key = $"{Subroutine.Name.NormalizedText}.{variableName}"
            End If

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

        Public Sub AddVariableInitialization(ByVal variable As TokenInfo)
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
