'#Form1{
'    TextBox1: TextBox
'    btnOk: Button
'    lblError: Label
'}

Public Class PreCompiler

    Private Shared ModuleInfo As New Dictionary(Of String, List(Of String))

    Shared Sub New()
        ListModuleMembers(GetType(Form))
        ListModuleMembers(GetType(Control))
        ListModuleMembers(GetType(TextBox))
        ListModuleMembers(GetType(Label))
        ListModuleMembers(GetType(Button))
        ListModuleMembers(GetType(ListBox))
    End Sub

    Private Shared Sub ListModuleMembers(t As Type)
        Dim members = (From m In t.GetMembers()
                       Select m.Name.ToLower).ToList

        ModuleInfo(t.Name) = members
    End Sub

    Public Shared Function GetModule(controlName As String, methodName As String) As String
        Dim method = methodName.ToLower
        If ModuleInfo(controlName).Contains(method) Then
            Return controlName
        ElseIf ModuleInfo(NameOf(Control)).Contains(method) Then
            Return NameOf(Control)
        Else
            Return ""
        End If
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