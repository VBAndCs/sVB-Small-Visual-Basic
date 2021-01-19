Imports System.Collections.Generic

Namespace Microsoft.SmallBasic.LanguageService
    Public Class CompletionItemDocumentation
        Private _paramsDoc As Dictionary(Of String, String)
        Public Property Example As String
        Public Property Returns As String
        Public Property Summary As String

        Public ReadOnly Property ParamsDoc As Dictionary(Of String, String)
            Get

                If _paramsDoc Is Nothing Then
                    _paramsDoc = New Dictionary(Of String, String)()
                End If

                Return _paramsDoc
            End Get
        End Property
    End Class
End Namespace
