Imports System.Reflection

Namespace Microsoft.SmallBasic.Completion
    Public Class CompletionItem
        Private replacementTextField As String
        Public Property Name As String
        Public Property DisplayName As String
        Public Property ItemType As CompletionItemType
        Public Property MemberInfo As MemberInfo

        Public Property ReplacementText As String
            Get

                If Not Equals(replacementTextField, Nothing) Then
                    Return replacementTextField
                End If

                Return DisplayName
            End Get
            Set(ByVal value As String)
                replacementTextField = value
            End Set
        End Property

        Public Overrides Function ToString() As String
            Return DisplayName
        End Function
    End Class
End Namespace
