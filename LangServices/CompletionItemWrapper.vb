Imports System
Imports System.Collections.Generic
Imports System.Reflection
Imports System.Text
Imports Microsoft.SmallBasic.Completion
Imports SmallBasicLibrary.Microsoft.SmallBasic.Library

Namespace Microsoft.SmallBasic.LanguageService
    Public Class CompletionItemWrapper
        Private _item As CompletionItem
        Private Shared _moduleDocMap As Dictionary(Of String, ModuleDocumentation) = New Dictionary(Of String, ModuleDocumentation)()
        Private _documentation As CompletionItemDocumentation

        Public ReadOnly Property CompletionItem As CompletionItem
            Get
                Return _item
            End Get
        End Property

        Public ReadOnly Property Name As String
            Get
                Return _item.DisplayName
            End Get
        End Property

        Public ReadOnly Property Description As String
            Get
                Return _item.MemberInfo.Name
            End Get
        End Property

        Public ReadOnly Property Prefix As String
            Get

                Select Case SymbolType
                    Case SymbolType.Keyword
                        Return "M:"
                    Case SymbolType.Type
                        Return "T:"
                    Case SymbolType.Method
                        Return "M:"
                    Case SymbolType.Property
                        Return "P:"
                    Case SymbolType.Event
                        Return "E:"
                    Case Else
                        Return "?:"
                End Select
            End Get
        End Property

        Public ReadOnly Property SymbolType As SymbolType
            Get

                Select Case _item.ItemType
                    Case CompletionItemType.EventName
                        Return SymbolType.Event
                    Case CompletionItemType.Identifier
                        Return SymbolType.Unknown
                    Case CompletionItemType.Keyword
                        Return SymbolType.Keyword
                    Case CompletionItemType.Label
                        Return SymbolType.Label
                    Case CompletionItemType.MethodName
                        Return SymbolType.Method
                    Case CompletionItemType.PropertyName
                        Return SymbolType.Property
                    Case CompletionItemType.SubroutineName
                        Return SymbolType.Subroutine
                    Case CompletionItemType.TypeName
                        Return SymbolType.Type
                    Case CompletionItemType.Variable
                        Return SymbolType.Variable
                    Case Else
                        Return SymbolType.Unknown
                End Select
            End Get
        End Property

        Public ReadOnly Property Display As String
            Get
                Return _item.DisplayName
            End Get
        End Property

        Public ReadOnly Property Documentation As CompletionItemDocumentation
            Get
                Return _documentation
            End Get
        End Property

        Public Sub New(ByVal item As CompletionItem)
            _item = item
            Dim normalizedModuleName As String = GetNormalizedModuleName()
            Dim value As ModuleDocumentation = Nothing

            If Not _moduleDocMap.TryGetValue(normalizedModuleName, value) Then
                value = New ModuleDocumentation(normalizedModuleName)
                _moduleDocMap(normalizedModuleName) = value
            End If

            _documentation = value.GetItemDocumentation(GetSymbolName())
        End Sub

        Private Function GetNormalizedModuleName() As String
            Dim text = If(Not _item.MemberInfo IsNot Nothing, GetType(Primitive).Module.FullyQualifiedName, _item.MemberInfo.Module.FullyQualifiedName)
            Return text.ToLowerInvariant()
        End Function

        Private Function GetSymbolName() As String
            Dim result = ""

            If SymbolType = SymbolType.Keyword Then
                result = $"{Prefix}Microsoft.SmallBasic.Library.Keywords.{_item.DisplayName}"
            ElseIf SymbolType = SymbolType.Method Then
                Dim methodInfo As MethodInfo = TryCast(_item.MemberInfo, MethodInfo)
                Dim parameters As ParameterInfo() = methodInfo?.GetParameters()

                If parameters?.Length > 0 Then
                    Dim stringBuilder As StringBuilder = New StringBuilder($"{Prefix}{methodInfo.DeclaringType.FullName}.{methodInfo.Name}")
                    stringBuilder.Append("(")
                    Dim array = parameters

                    For Each parameterInfo In array
                        stringBuilder.AppendFormat("{0},", parameterInfo.ParameterType.FullName)
                    Next

                    stringBuilder.Remove(stringBuilder.Length - 1, 1)
                    stringBuilder.Append(")")
                    result = stringBuilder.ToString()
                Else
                    result = $"{Prefix}{methodInfo?.DeclaringType.FullName}.{methodInfo?.Name}"
                End If
            ElseIf SymbolType = SymbolType.Property Then
                Dim propertyInfo As PropertyInfo = TryCast(_item.MemberInfo, PropertyInfo)
                result = $"{Prefix}{propertyInfo?.DeclaringType.FullName}.{propertyInfo?.Name}"
            ElseIf SymbolType = SymbolType.Event Then
                Dim eventInfo As EventInfo = TryCast(_item.MemberInfo, EventInfo)
                result = $"{Prefix}{eventInfo.DeclaringType.FullName}.{eventInfo.Name}"
            ElseIf SymbolType = SymbolType.Type Then
                Dim type As Type = TryCast(_item.MemberInfo, Type)
                result = $"{Prefix}{type.FullName}"
            End If

            Return result
        End Function
    End Class
End Namespace
