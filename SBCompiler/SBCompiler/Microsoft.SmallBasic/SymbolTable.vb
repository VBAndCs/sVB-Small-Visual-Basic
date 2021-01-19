Imports System.Collections.Generic
Imports System.Globalization

Namespace Microsoft.SmallBasic
    Public Class SymbolTable
        Private _errors As List(Of [Error])
        Private _initializedVariables As Dictionary(Of String, TokenInfo) = New Dictionary(Of String, TokenInfo)()
        Private _variables As Dictionary(Of String, TokenInfo) = New Dictionary(Of String, TokenInfo)()
        Private _subroutines As Dictionary(Of String, TokenInfo) = New Dictionary(Of String, TokenInfo)()
        Private _labels As Dictionary(Of String, TokenInfo) = New Dictionary(Of String, TokenInfo)()

        Public ReadOnly Property Errors As List(Of [Error])
            Get
                Return _errors
            End Get
        End Property

        Public ReadOnly Property InitializedVariables As Dictionary(Of String, TokenInfo)
            Get
                Return _initializedVariables
            End Get
        End Property

        Public ReadOnly Property Variables As Dictionary(Of String, TokenInfo)
            Get
                Return _variables
            End Get
        End Property

        Public ReadOnly Property Subroutines As Dictionary(Of String, TokenInfo)
            Get
                Return _subroutines
            End Get
        End Property

        Public ReadOnly Property Labels As Dictionary(Of String, TokenInfo)
            Get
                Return _labels
            End Get
        End Property

        Public Sub New(ByVal errors As List(Of [Error]))
            _errors = errors

            If _errors Is Nothing Then
                _errors = New List(Of [Error])()
            End If
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

        Public Sub AddSubroutine(ByVal subroutineName As TokenInfo)
            Dim normalizedText = subroutineName.NormalizedText

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
