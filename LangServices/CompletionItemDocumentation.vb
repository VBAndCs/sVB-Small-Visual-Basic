Imports System.Collections.Generic

Namespace Microsoft.SmallVisualBasic.LanguageService
    Public Class CompletionItemDocumentation
        Private _paramsDoc As Dictionary(Of String, String)
        Public Prefix As String
        Public Suffix As String
        Public Example As String
        Public Returns As String
        Public Summary As String

        Public Property ParamsDoc As Dictionary(Of String, String)
            Get
                If _paramsDoc Is Nothing Then
                    _paramsDoc = New Dictionary(Of String, String)()
                End If

                Return _paramsDoc
            End Get

            Set(value As Dictionary(Of String, String))
                _paramsDoc = value
            End Set
        End Property
    End Class
End Namespace
