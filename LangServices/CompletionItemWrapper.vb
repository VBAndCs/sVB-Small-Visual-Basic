Imports System.Reflection
Imports System.Text
Imports Microsoft.SmallVisualBasic.Completion
Imports Microsoft.SmallVisualBasic.Library

Namespace Microsoft.SmallVisualBasic.LanguageService
    Public Class CompletionItemWrapper
        Private _item As CompletionItem
        Private _enumName As String
        Private Shared _moduleDocMap As New Dictionary(Of String, ModuleDocumentation)()
        Private _docs As CompletionItemDocumentation

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
                Return _docs?.Summary
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
        Private isGlobal As Boolean

        Public NavigateTo As NavigateTo

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
                            If CompletionItem.ObjectName.ToLower() = "global" Then
                                _symbolType = SymbolType.Subroutine
                                isGlobal = True
                            Else
                                _symbolType = SymbolType.Method
                            End If


                        Case CompletionItemType.PropertyName
                            If CompletionItem.ObjectName.ToLower() = "global" Then
                                _symbolType = SymbolType.GlobalVariable
                                isGlobal = True
                            Else
                                _symbolType = SymbolType.Property
                            End If

                        Case CompletionItemType.SubroutineName
                            _symbolType = SymbolType.Subroutine

                        Case CompletionItemType.TypeName
                            If CompletionItem.Key = "global" Then
                                _symbolType = SymbolType.GlobalModule
                            Else
                                _symbolType = SymbolType.Type
                            End If

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
                If _enumName = "" Then
                    Return _item.DisplayName
                Else
                    Return _enumName & "." & _item.DisplayName
                End If
            End Get
        End Property

        Public ReadOnly Property Documentation As CompletionItemDocumentation
            Get
                Return _docs
            End Get
        End Property

        Public ReadOnly Property ReplacementText As String
            Get
                If _item.ReplacementText = "" Then
                    Return Display
                ElseIf _enumName = "" Then
                    Return _item.ReplacementText
                Else
                    Return _enumName & "." & _item.ReplacementText
                End If
            End Get
        End Property

        Public Sub New(item As CompletionItem, bag As CompletionBag)
            _item = item
            Dim enumName = bag.SelectEspecialItem
            _enumName = If(enumName = "*", "", enumName)

            Dim moduleName = GetNormalizedModuleName()
            Dim moduleDoc As ModuleDocumentation = Nothing

            If moduleName <> "" AndAlso Not _moduleDocMap.TryGetValue(moduleName, moduleDoc) Then
                moduleDoc = New ModuleDocumentation(moduleName)
                _moduleDocMap(moduleName) = moduleDoc
            End If

            Select Case SymbolType
                Case SymbolType.Subroutine
                    Dim parseTree = If(isGlobal, bag.GlobalParseTree, bag.ParseTree)
                    If parseTree Is Nothing Then Return

                    Dim subrotine = CompletionHelper.GetSubroutine(_item.DisplayName, parseTree)
                    If subrotine Is Nothing Then Return

                    If isGlobal Then
                        item.DefinitionIdintifier = subrotine.Name
                        NavigateTo = NavigateTo.GlobalModule
                    End If

                    _docs = New CompletionItemDocumentation() With {
                            .Prefix = subrotine.SubToken.Text & " ",
                            .Summary = subrotine.GetSummery,
                            .Suffix = InferType(
                                   subrotine.Name.LCaseText,
                                   If(isGlobal, bag.GlobalSymbolTable, bag.SymbolTable)
                             ),
                            .ParamsDoc = subrotine.GetParamsDoc(),
                            .Returns = subrotine.GetRetunDoc()
                    }

                Case SymbolType.DynamicProperty
                    _docs = New CompletionItemDocumentation() With {
                            .Prefix = "Dynamic Property: " & item.ObjectName & "!",
                            .Suffix = InferType(item.Key, bag.SymbolTable),
                            .Summary = item.DefinitionIdintifier.Comment
                    }

                Case SymbolType.Label
                    _docs = New CompletionItemDocumentation() With {
                            .Prefix = "Label: ",
                            .Summary = item.DefinitionIdintifier.Comment
                    }

                Case SymbolType.Literal
                    Dim result = Parser.ParseDateLiteral(item.DisplayName)
                    _docs = New CompletionItemDocumentation() With {
                            .Prefix = If(result.IsDate, "Date Literal: ", "Time Span Literal: "),
                            .Summary = If(Not result.Ticks.HasValue,
                                    $"Invalid {If(result.IsDate, "date", "time span")} format!",
                                    "Value = """ & If(result.IsDate,
                                            New Date(result.Ticks.Value).ToString("MM/dd/yyyy hh:mm:ss.FFFFFFF tt"),
                                            FormatTimeSpan(result.Ticks.Value)
                                    ) & """"
                            )
                    }

                Case SymbolType.GlobalVariable
                    Dim symbolTable As SymbolTable
                    If isGlobal Then
                        symbolTable = bag.GlobalSymbolTable
                        NavigateTo = NavigateTo.GlobalModule
                        If symbolTable IsNot Nothing Then
                            item.DefinitionIdintifier = symbolTable.GlobalVariables(item.Key)
                        End If

                    Else
                        symbolTable = bag.SymbolTable
                    End If

                    _docs = New CompletionItemDocumentation() With {
                            .Prefix = "Global Variable: ",
                            .Suffix = If(
                                    isGlobal OrElse item.ObjectName = "",
                                    InferType(item.Key, symbolTable),
                                    $" As {item.ObjectName}"
                            ),
                            .Summary = item.DefinitionIdintifier.Comment
                    }

                Case SymbolType.LocalVariable
                    Dim vars = bag.SymbolTable.LocalVariables
                    Dim var = item.Key
                    If Not vars.ContainsKey(var) Then Return
                    Dim varExpr = vars(var)

                    _docs = New CompletionItemDocumentation() With {
                            .Prefix = If(varExpr.IsParam,
                                "Parameter: " & varExpr.Subroutine.Name.Text & ".",
                                "Local Variable: "
                            ),
                            .Suffix = If(
                                item.ObjectName = "" OrElse item.ObjectName = "Forms",
                                InferType(item.Key, bag.SymbolTable),
                                $" As {item.ObjectName}"
                            ),
                            .Summary = varExpr.Identifier.Comment
                    }

                Case SymbolType.GlobalModule
                    item.DefinitionIdintifier = New Token() With {
                        .Column = -1,
                        .Type = TokenType.Identifier
                    }
                    NavigateTo = NavigateTo.GlobalModule

                    _docs = New CompletionItemDocumentation() With {
                            .Suffix = " Type",
                            .Summary = "The global module of the project. You can define global functions and variables in this module and use it from any form in the project."
                    }

                Case SymbolType.Control
                    NavigateTo = NavigateTo.Designer
                    _docs = New CompletionItemDocumentation() With {
                            .Prefix = "Global Variable: ",
                            .Suffix = If(item.ObjectName = "", "", $" As {item.ObjectName}"),
                            .Summary = If(item.DisplayName = "Me",
                                 $"Me is a global variable that referes to the current form, which is {item.Key} in this context",
                                 $"A global variable that referes to a {item.ObjectName} control that you created by the form designer"
                            )
                    }

                Case Else
                    Dim name = GetSymbolName()
                    If name <> "" AndAlso moduleDoc IsNot Nothing Then
                        _docs = moduleDoc.GetItemDocumentation(name)
                        If _docs Is Nothing Then
                            If item.MemberInfo IsNot Nothing AndAlso item.MemberInfo.MemberType = MemberTypes.Method Then
                                Dim m = CType(item.MemberInfo, MethodInfo)
                                Dim paramsDocs As New Dictionary(Of String, String)
                                Dim params = m.GetParameters()
                                Dim isExMethod = m.GetCustomAttributes(GetType(WinForms.ExMethodAttribute), inherit:=False).Count > 0

                                If params IsNot Nothing Then
                                    For Each param In params
                                        If isExMethod Then
                                            isExMethod = False
                                            Continue For
                                        End If
                                        paramsDocs(param.Name) = ""
                                    Next
                                End If

                                _docs = New CompletionItemDocumentation()
                                _docs.ParamsDoc = paramsDocs
                            End If
                        End If

                        If _docs IsNot Nothing Then _docs.Suffix = InferType(item.MemberInfo)
                    End If
            End Select


        End Sub

        Private Function FormatTimeSpan(ticks As Long) As String
            Dim ts = New TimeSpan(ticks).ToString()
            Dim pos = ts.LastIndexOf(":")
            If pos = -1 Then Return ts

            pos = ts.IndexOf(".", pos)
            If pos = -1 Then Return ts

            ts = ts.TrimEnd("0"c)
            If ts.EndsWith(".") Then ts = ts.Substring(0, ts.Length - 1)

            Return ts
        End Function

        Private Function InferType(memberInfo As MemberInfo) As String
            If memberInfo Is Nothing Then Return ""

            Dim attrs = memberInfo.GetCustomAttributes(
                GetType(WinForms.ReturnValueTypeAttribute),
                False
            )

            If attrs Is Nothing OrElse attrs.Count = 0 Then Return ""
            Dim type = CType(
                attrs(0),
                WinForms.ReturnValueTypeAttribute
            ).ReturnTypeValue.ToString()
            Return " As " & type
        End Function

        Private Function InferType(key As String, symbolTable As SymbolTable) As String
            Dim varType = symbolTable.GetInferedType(key.ToLower())
            If varType <> VariableType.Any Then
                Return " As " & varType.ToString
            Else
                Return ""
            End If
        End Function

        Private Function GetNormalizedModuleName() As String
            Dim text = If(_item.MemberInfo Is Nothing,
                GetType(Primitive).Module.FullyQualifiedName,
                _item.MemberInfo.Module.FullyQualifiedName
            ).ToLowerInvariant()

            Return If(text.StartsWith("<"), "", text)
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
                Dim methodInfo = TryCast(_item.MemberInfo, MethodInfo)
                Dim parameters = methodInfo?.GetParameters()

                If parameters?.Length > 0 Then
                    Dim stringBuilder As New StringBuilder($"M:{methodInfo.DeclaringType.FullName}.{methodInfo.Name}")
                    stringBuilder.Append("(")

                    For Each parameterInfo In parameters
                        stringBuilder.Append(parameterInfo.ParameterType.FullName)
                        stringBuilder.Append(",")
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
