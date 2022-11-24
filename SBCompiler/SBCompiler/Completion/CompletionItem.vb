Imports System.Reflection

Namespace Microsoft.SmallVisualBasic.Completion
    Public Class CompletionItem

        Private _replacementText As String

        Public ParamIndex As Integer = -1
        Public ObjectName As String
        Public Key As String
        Public Property DisplayName As String
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

        Public Function GetHistoryKey(Optional prefix As String = "") As String
            Select Case ItemType
                Case CompletionItemType.EventName,
                         CompletionItemType.MethodName,
                         CompletionItemType.PropertyName,
                         CompletionItemType.DynamicPropertyName

                    Dim k As String = ""
                    If MemberInfo?.DeclaringType IsNot Nothing Then
                        k = MemberInfo.DeclaringType.Name.ToLower()
                        If k = "control" Then
                            k = WinForms.PreCompiler.GetModuleFromVarName(ObjectName)
                        End If
                    End If

                    If k = "" Then k = ObjectName
                    Return k?.ToLower()

                Case CompletionItemType.GlobalVariable,
                          CompletionItemType.LocalVariable,
                          CompletionItemType.SubroutineName

                    If prefix <> "" Then Return "_" & prefix.ToLower()(0)
                    Return "_" & DisplayName.ToLower()(0)

                Case Else
                    If prefix <> "" Then Return "_" & prefix.ToLower()(0)
                    If DisplayName = "" Then Return ""

                    Dim x = DisplayName.ToLower()
                    If x = "me" Then Return "_m"
                    Return "_t_" & x(0)
            End Select
        End Function

        Public Overrides Function ToString() As String
            Return DisplayName
        End Function
    End Class
End Namespace
