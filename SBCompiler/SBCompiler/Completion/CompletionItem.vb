Imports System.Reflection

Namespace Microsoft.SmallVisualBasic.Completion
    Public Class CompletionItem

        Private _replacementText As String

        Public ParamIndex As Integer = -1
        Public ObjectName As String
        Public Key As String
        Public DisplayName As String
        Public ItemType As CompletionItemType
        Public Property MemberInfo As MemberInfo
        Public DefinitionIdintifier As Token

        Public Property ReplacementText As String
            Get
                If Not Equals(_replacementText, Nothing) Then
                    Return _replacementText
                End If

                Return DisplayName
            End Get

            Set(value As String)
                _replacementText = value
            End Set
        End Property

        Public ReadOnly Property HistoryKey As String
            Get
                Select Case ItemType
                    Case CompletionItemType.EventName,
                             CompletionItemType.MethodName,
                             CompletionItemType.PropertyName,
                             CompletionItemType.DynamicPropertyName

                        If ObjectName <> "" Then
                            Return ObjectName.ToLower()
                        Else
                            Return MemberInfo?.DeclaringType?.Name.ToLower()
                        End If

                    Case CompletionItemType.GlobalVariable,
                              CompletionItemType.LocalVariable,
                              CompletionItemType.SubroutineName
                        Return "_" & DisplayName.ToLower()(0)

                    Case Else
                        Return ""
                End Select
            End Get
        End Property

        Public Overrides Function ToString() As String
            Return DisplayName
        End Function
    End Class
End Namespace
