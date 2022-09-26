Imports System
Imports System.Collections.Generic
Imports System.Reflection
Imports System.Text
Imports Microsoft.SmallVisualBasic.Completion
Imports Microsoft.SmallVisualBasic.Library

Namespace Microsoft.SmallVisualBasic.LanguageService
    Public Class CompletionItemWrapper
        Private _item As CompletionItem
        Private Shared _moduleDocMap As New Dictionary(Of String, ModuleDocumentation)()
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

        Public ReadOnly Property Summary As String
            Get
                Return _documentation?.Summary
            End Get
        End Property

        Public ReadOnly Property Description As String
            Get
                Return _item.MemberInfo?.Name
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
                        Return ""
                End Select
            End Get
        End Property

        Dim _symbolType? As SymbolType

        Public ReadOnly Property SymbolType As SymbolType
            Get
                If Not _symbolType.HasValue Then

                    Select Case _item.ItemType
                        Case CompletionItemType.EventName
                            _symbolType = SymbolType.Event

                        Case CompletionItemType.Identifier
                            _symbolType = SymbolType.Unknown

                        Case CompletionItemType.Keyword
                            _symbolType = SymbolType.Keyword

                        Case CompletionItemType.Label
                            _symbolType = SymbolType.Label

                        Case CompletionItemType.MethodName
                            _symbolType = SymbolType.Method

                        Case CompletionItemType.PropertyName
                            _symbolType = SymbolType.Property

                        Case CompletionItemType.SubroutineName
                            _symbolType = SymbolType.Subroutine

                        Case CompletionItemType.TypeName
                            _symbolType = SymbolType.Type

                        Case CompletionItemType.Control
                            _symbolType = SymbolType.Control

                        Case CompletionItemType.GlobalVariable
                            _symbolType = SymbolType.GlobalVariable

                        Case CompletionItemType.LocalVariable
                            _symbolType = SymbolType.LocalVariable

                        Case CompletionItemType.DynamicPropertyName
                            _symbolType = SymbolType.DynamicProperty

                        Case CompletionItemType.Lireral
                            _symbolType = SymbolType.Literal

                        Case Else
                            _symbolType = SymbolType.Unknown
                    End Select
                End If

                Return _symbolType.Value
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

        Public Sub New(item As CompletionItem, bag As CompletionBag)
            _item = item
            Dim normalizedModuleName = GetNormalizedModuleName()
            Dim value As ModuleDocumentation = Nothing

            If Not _moduleDocMap.TryGetValue(normalizedModuleName, value) Then
                value = New ModuleDocumentation(normalizedModuleName)
                _moduleDocMap(normalizedModuleName) = value
            End If

            Select Case SymbolType
                Case SymbolType.Subroutine
                    Dim parseTree = bag.ParseTree
                    If parseTree Is Nothing Then Return

                    Dim subrotine = CompletionHelper.GetSubroutine(_item.DisplayName, parseTree)
                    If subrotine Is Nothing Then Return

                    _documentation = New CompletionItemDocumentation() With {
                            .Prefix = subrotine.StartToken.Text & " ",
                            .Summary = subrotine.GetSummery,
                            .ParamsDoc = subrotine.GetParamsDoc(),
                            .Returns = subrotine.GetRetunDoc()
                    }

                Case SymbolType.DynamicProperty
                    _documentation = New CompletionItemDocumentation() With {
                            .Prefix = item.ObjectName & "!",
                            .Suffix = " Dynamic Property",
                            .Summary = item.DefinitionIdintifier.Comment
                    }

                Case SymbolType.Label
                    _documentation = New CompletionItemDocumentation() With {
                            .Prefix = "Label: ",
                            .Summary = item.DefinitionIdintifier.Comment
                    }

                Case SymbolType.Literal
                    Dim result = Parser.ParseDateLiteral(item.DisplayName)
                    _documentation = New CompletionItemDocumentation() With {
                            .Prefix = If(result.IsDate, "Date Literal: ", "Time Span Literal: "),
                            .Summary = If(result.Ticks = 0,
                                    $"Invalid {If(result.IsDate, "date", "time span")} format!",
                                    "Value = """ & If(result.IsDate,
                                            New Date(result.Ticks),
                                            New TimeSpan(result.Ticks)
                                    ).ToString() & """"
                            )
                    }

                Case SymbolType.GlobalVariable
                    _documentation = New CompletionItemDocumentation() With {
                            .Prefix = "Global Variable: ",
                            .Suffix = If(item.ObjectName = "", InferType(item.Key, bag), $" As {item.ObjectName}"),
                            .Summary = item.DefinitionIdintifier.Comment
                    }

                Case SymbolType.LocalVariable
                    Dim vars = bag.SymbolTable.LocalVariables
                    Dim var = item.Key
                    If Not vars.ContainsKey(var) Then Return
                    Dim varExpr = vars(var)

                    _documentation = New CompletionItemDocumentation() With {
                            .Prefix = If(varExpr.IsParam, "Parameter: " & varExpr.Subroutine.Name.Text & ".", "Local Variable: "),
                            .Suffix = If(item.ObjectName = "" OrElse item.ObjectName = "Forms", InferType(item.Key, bag), $" As {item.ObjectName}"),
                            .Summary = varExpr.Identifier.Comment
                    }

                Case SymbolType.Control
                    _documentation = New CompletionItemDocumentation() With {
                            .Prefix = "Global Variable: ",
                            .Suffix = If(item.ObjectName = "", "", $" As {item.ObjectName}"),
                            .Summary = If(item.DisplayName = "Me",
                                 $"Me is a global variable that referes to the current form, which is {item.Key} in this context",
                                 $"A global variable that referes to a {item.ObjectName} control that you created by the form designer"
                            )
                    }

                Case Else
                    Dim name = GetSymbolName()
                    If name <> "" Then
                        _documentation = value.GetItemDocumentation(name)
                    End If
            End Select


        End Sub

        Private Function InferType(key As String, bag As CompletionBag) As String
            Dim varType = bag.SymbolTable.GetInferedType(key)
            If varType <> VariableType.None Then
                Return " As " & varType.ToString
            Else
                Return ""
            End If
        End Function

        Private Function GetNormalizedModuleName() As String
            Dim text = If(Not _item.MemberInfo IsNot Nothing, GetType(Primitive).Module.FullyQualifiedName, _item.MemberInfo.Module.FullyQualifiedName)
            Return text.ToLowerInvariant()
        End Function

        Private Function GetSymbolName() As String
            Dim result = ""

            Select Case SymbolType
                Case SymbolType.Keyword
                    result = $"{Prefix}Microsoft.SmallVisualBasic.Library.Keywords.{_item.DisplayName}"

                Case SymbolType.Method
                    result = GetMethodName()

                Case SymbolType.Property
                    Dim propertyInfo As PropertyInfo = TryCast(_item.MemberInfo, PropertyInfo)
                    If propertyInfo Is Nothing Then
                        result = GetMethodName()
                    Else
                        result = $"{Prefix}{propertyInfo.DeclaringType.FullName}.{propertyInfo.Name}"
                    End If

                Case SymbolType.Event
                    Dim eventInfo = TryCast(_item.MemberInfo, EventInfo)
                    result = $"{Prefix}{eventInfo.DeclaringType.FullName}.{eventInfo.Name}"

                Case SymbolType.Type
                    Dim type = TryCast(_item.MemberInfo, Type)
                    result = $"{Prefix}{type.FullName}"
            End Select

            Return result
        End Function

        Private Function GetMethodName() As String
            Dim result = ""
            Try
                Dim methodInfo As MethodInfo = TryCast(_item.MemberInfo, MethodInfo)
                Dim parameters As ParameterInfo() = methodInfo?.GetParameters()

                If parameters?.Length > 0 Then
                    Dim stringBuilder As StringBuilder = New StringBuilder($"M:{methodInfo.DeclaringType.FullName}.{methodInfo.Name}")
                    stringBuilder.Append("(")
                    Dim array = parameters

                    For Each parameterInfo In array
                        stringBuilder.AppendFormat("{0},", parameterInfo.ParameterType.FullName)
                    Next

                    stringBuilder.Remove(stringBuilder.Length - 1, 1)
                    stringBuilder.Append(")")
                    result = stringBuilder.ToString()
                Else
                    result = $"M:{methodInfo?.DeclaringType.FullName}.{methodInfo?.Name}"
                End If

            Catch ex As Exception
            End Try

            Return result
        End Function
    End Class
End Namespace
