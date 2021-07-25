Imports System.Collections.Generic
Imports System.Globalization

Namespace Microsoft.SmallBasic
    Public Class SymbolTable

        Public ReadOnly Property Errors As List(Of [Error])

        Public ReadOnly Property InitializedVariables As New Dictionary(Of String, TokenInfo)

        Public ReadOnly Property Variables As New Dictionary(Of String, TokenInfo)

        Public ReadOnly Property Subroutines As New Dictionary(Of String, TokenInfo)

        Public ReadOnly Property Labels As New Dictionary(Of String, TokenInfo)

        Public Sub New(ByVal errors As List(Of [Error]))
            _errors = errors

            If _errors Is Nothing Then
                _errors = New List(Of [Error])()
            End If
        End Sub

        Public Sub CopyFrom(symbolTable As SymbolTable)
            Copy(symbolTable.Variables, Me.Variables)
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
            _variables.Clear()
        End Sub

        Public Sub AddVariable(ByVal variable As TokenInfo)
            If Not Variables.ContainsKey(variable.NormalizedText) Then
                Variables.Add(variable.NormalizedText, variable)
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

            If Variables.ContainsKey(normalizedText) Then
                Variables.Remove(normalizedText)
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
