'#Form1{
'    TextBox1: TextBox
'    btnOk: Button
'    lblError: Label
'}

Imports Microsoft.SmallBasic.Completion

Public NotInheritable Class PreCompiler

    Private Shared ModuleInfo As New Dictionary(Of String, List(Of String))
    Private Shared CompletionItems As New Dictionary(Of String, List(Of CompletionItem))

    Shared Sub New()
        ListModuleMembers(GetType(Forms))
        ListModuleMembers(GetType(Form))
        ListModuleMembers(GetType(Control))
        ListModuleMembers(GetType(TextBox))
        ListModuleMembers(GetType(Label))
        ListModuleMembers(GetType(Button))
        ListModuleMembers(GetType(ListBox))

        AddCompletionList(GetType(Form))
        AddCompletionList(GetType(Control))
        AddCompletionList(GetType(TextBox))
        AddCompletionList(GetType(Label))
        AddCompletionList(GetType(Button))
        AddCompletionList(GetType(ListBox))
    End Sub

    Public Function GetBaseTypes(name As String) As Type()
        Select Case name
            Case NameOf(Form)
                Return {GetType(Form), GetType(Forms), GetType(Control)}
            Case NameOf(Control)
                Return {GetType(Control)}
            Case Else
                Dim t = Type.GetType("SmallBasic.WinForms." & name)
                If t Is Nothing Then
                    Return {GetType(Control)}
                Else
                    Return {GetType(Form), GetType(Control)}
                End If
        End Select
    End Function

    Dim PrimativeType = GetType(Microsoft.SmallBasic.Library.Primitive)

    Public Function GetExtenstions(forType As Type, inType As Type) As List(Of Reflection.MemberInfo)
        If forType.Name = NameOf(Form) Then
            Return (From M In inType.GetMethods(Reflection.BindingFlags.Public)
                    Where M.GetParameters().Any(Function(p) p.ParameterType Is PrimativeType AndAlso (p.Name = "formName"))
                    Select CType(M, Reflection.MemberInfo)).ToList
        Else
            Return (From M In inType.GetMethods(Reflection.BindingFlags.Public)
                    Where M.GetParameters().Any(Function(p) p.ParameterType Is PrimativeType AndAlso (p.Name = "controlName"))
                    Select CType(M, Reflection.MemberInfo)).ToList
        End If
    End Function

    Private Shared Sub ListModuleMembers(t As Type)
        Dim members = (From m In t.GetMembers()
                       Select m.Name.ToLower).ToList

        ModuleInfo(t.Name) = members
    End Sub

    Public Shared Function GetMethodInfo(controlName As String, methodName As String) As ([Module] As String, ParamsCount As Integer)
        Dim method = methodName.ToLower
        Dim moduleName = ""
        If ModuleInfo(controlName).Contains(method) Then
            moduleName = controlName
        ElseIf ModuleInfo(NameOf(Control)).Contains(method) Then
            moduleName = NameOf(Control)
        ElseIf ModuleInfo(NameOf(Forms)).Contains(method) Then
            moduleName = NameOf(Forms)
        Else
            Return ("", 0)
        End If

        Dim t = Type.GetType("SmallBasic.WinForms." & moduleName)
        Dim params = t?.GetMethod(methodName)?.GetParameters()
        Return (moduleName, If(params Is Nothing, 0, params.Length))
    End Function

    Public Shared Function ParseFormHints(txt As String) As FormInfo
        If Not txt.StartsWith("'#") Then Return Nothing

        Dim pos1 = 2
        Dim pos2 = txt.IndexOf(Environment.NewLine, pos1)
        If pos2 = -1 Then Return Nothing
        If txt(pos2 - 1) <> "{" Then Return Nothing

        Dim info As New FormInfo
        info.Form = txt.Substring(pos1, pos2 - 1 - pos1).Trim()
        If info.Form = "" Then Return Nothing
        info.ControlsInfo.Add(info.Form, "Form")

        Do
            pos1 = pos2 + Environment.NewLine.Length
            If txt.Substring(pos1, 5) = "'    " Then
                pos1 += 5
                pos2 = txt.IndexOf(Environment.NewLine, pos1)
                If pos2 = -1 Then Return Nothing
                Dim hint = txt.Substring(pos1, pos2 - pos1)
                Dim pair = hint.Split(":"c)
                If pair Is Nothing OrElse pair.Length <> 2 Then Return Nothing
                info.ControlsInfo.Add(pair(0).Trim, pair(1).Trim)
            ElseIf txt.Substring(pos1, 2) = "'}" Then
                Return info
            Else
                Return Nothing
            End If
        Loop
    End Function


    Private Shared Sub AddCompletionList(type As Type)
        Dim compList As New List(Of CompletionItem)

        Dim methods = type.GetMethods(Reflection.BindingFlags.Static Or Reflection.BindingFlags.Public)
        Dim extensionParams = If(type.Name = "Form", 1, 2)

        For Each methodInfo In methods
            Dim name = ""
            Dim completionItem As New CompletionItem()
            If methodInfo.GetCustomAttributes(GetType(ExMethodAttribute), inherit:=False).Count > 0 Then
                name = methodInfo.Name
                completionItem.Name = name
                completionItem.DisplayName = name
                completionItem.ItemType = CompletionItemType.MethodName
                If methodInfo.GetParameters().Length > extensionParams Then
                    completionItem.ReplacementText = name & "("
                Else
                    completionItem.ReplacementText = name & "()"
                End If
                compList.Add(completionItem)

            ElseIf methodInfo.Name.ToLower().StartsWith("get") AndAlso methodInfo.GetCustomAttributes(GetType(ExPropertyAttribute), inherit:=False).Count > 0 Then
                name = methodInfo.Name.Substring(3)
                completionItem.Name = name
                completionItem.DisplayName = name
                completionItem.ItemType = CompletionItemType.PropertyName
                completionItem.ReplacementText = name
                compList.Add(completionItem)
            End If
        Next

        Dim events = type.GetEvents(Reflection.BindingFlags.Static Or Reflection.BindingFlags.Public)
        For Each eventInfo In events
            If eventInfo.EventHandlerType Is GetType(Microsoft.SmallBasic.Library.SmallBasicCallback) Then
                Dim name = eventInfo.Name
                compList.Add(New CompletionItem() With {
                    .Name = name,
                    .DisplayName = name,
                    .ItemType = CompletionItemType.MethodName,
                    .ReplacementText = name
                })
            End If
        Next

        CompletionItems.Add(type.Name, compList)

    End Sub

    Public Shared Sub FillMemberNames(completionBag As CompletionBag, moduleName As String)
        completionBag.CompletionItems.AddRange(CompletionItems("Control"))

        If moduleName <> "Control" Then
            completionBag.CompletionItems.AddRange(CompletionItems(moduleName))
        End If

    End Sub
End Class

Public Class FormInfo
    Public Form As String
    Public ControlsInfo As New Dictionary(Of String, String)
End Class


Public Class ExPropertyAttribute
    Inherits Attribute

End Class

Public Class ExMethodAttribute
    Inherits Attribute

End Class

Public Class ExEventAttribute
    Inherits Attribute

End Class
