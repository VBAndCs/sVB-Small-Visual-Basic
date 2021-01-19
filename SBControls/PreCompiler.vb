'#Form1{
'    TextBox1: TextBox
'    btnOk: Button
'    lblError: Label
'}

Public NotInheritable Class PreCompiler

    Private Shared ModuleInfo As New Dictionary(Of String, List(Of String))

    Shared Sub New()
        ListModuleMembers(GetType(Forms))
        ListModuleMembers(GetType(Form))
        ListModuleMembers(GetType(Control))
        ListModuleMembers(GetType(TextBox))
        ListModuleMembers(GetType(Label))
        ListModuleMembers(GetType(Button))
        ListModuleMembers(GetType(ListBox))
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
                    From p In M.GetParameters()
                    Where p.ParameterType Is PrimativeType AndAlso (p.Name = "formName")
                    Select CType(M, Reflection.MemberInfo)).ToList
        Else
            Return (From M In inType.GetMethods(Reflection.BindingFlags.Public)
                    From p In M.GetParameters()
                    Where p.ParameterType Is PrimativeType AndAlso (p.Name = "controlName")
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

End Class

Public Class FormInfo
    Public Form As String
    Public ControlsInfo As New Dictionary(Of String, String)
End Class