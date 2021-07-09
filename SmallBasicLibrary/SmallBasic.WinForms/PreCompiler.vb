

Imports Microsoft.SmallBasic.Library

Namespace WinForms
    Public NotInheritable Class PreCompiler

        Private Shared ModuleInfo As New Dictionary(Of String, List(Of String))
        Private Shared EventsInfo As New Dictionary(Of String, List(Of String))

        Shared Sub New()
            ListModuleMembers(GetType(Forms))
            ListModuleMembers(GetType(Form))
            ListModuleMembers(GetType(Control))
            ListModuleMembers(GetType(TextBox))
            ListModuleMembers(GetType(Label))
            ListModuleMembers(GetType(Button))
            ListModuleMembers(GetType(ListBox))

        End Sub

        Public Shared Function GetBaseTypes(name As String) As Type()
            Select Case name
                Case NameOf(Form)
                    Return {GetType(Form), GetType(Forms), GetType(Control)}
                Case NameOf(Control)
                    Return {GetType(Control)}
                Case Else
                    Dim t = Type.GetType(WinFormsNS & name)
                    If t Is Nothing Then
                        Return {GetType(Control)}
                    Else
                        Return {GetType(Form), GetType(Control)}
                    End If
            End Select
        End Function

        Dim PrimativeType As Type = GetType(Primitive)
        Private Const WinFormsNS As String = "Microsoft.SmallBasic.WinForms."

        Private Shared Sub ListModuleMembers(t As Type)
            Dim members = (From m In t.GetMembers()
                           Select m.Name.ToLower).ToList

            ModuleInfo(t.Name) = members

            Dim events = (From e In t.GetEvents()
                          Where e.EventHandlerType Is GetType(SmallBasicCallback)
                          Select e.Name).ToList()

            EventsInfo(t.Name) = events
        End Sub

        Public Shared Function GetEvents(controlName As String) As List(Of String)
            Dim events As New List(Of String)
            For Each e In EventsInfo(NameOf(Control))
                events.Add(e)
            Next
            If controlName = NameOf(Control) Then Return events

            For Each e In EventsInfo(controlName)
                If Not events.Contains(e) Then events.Add(e)
            Next

            Return events
        End Function

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

            Dim t = Type.GetType(WinFormsNS & moduleName)
            Dim params = t?.GetMethod(methodName)?.GetParameters()
            Return (moduleName, If(params Is Nothing, 0, params.Length))
        End Function

        Public Shared Function ParseFormHints(txt As String) As FormInfo
            '#Form1{
            '    TextBox1: TextBox
            '    btnOk: Button
            '    lblError: Label
            '}

            If Not txt.StartsWith("'#") Then Return Nothing

            Dim pos1 = 2
            Dim pos2 = txt.IndexOf(Environment.NewLine, pos1)
            If pos2 = -1 Then Return Nothing
            If txt(pos2 - 1) <> "{" Then Return Nothing

            Dim info As New FormInfo
            info.Form = txt.Substring(pos1, pos2 - 1 - pos1).Trim()
            If info.Form = "" Then Return Nothing
            info.ControlsInfo.Add(info.Form.ToLower(), "Form")
            info.ControlNames.Add(info.Form)

            Do
                pos1 = pos2 + Environment.NewLine.Length
                If txt.Substring(pos1, 5) = "'    " Then
                    pos1 += 5
                    pos2 = txt.IndexOf(Environment.NewLine, pos1)
                    If pos2 = -1 Then Return Nothing
                    Dim hint = txt.Substring(pos1, pos2 - pos1)
                    Dim pair = hint.Split(":"c)
                    If pair Is Nothing OrElse pair.Length <> 2 Then Return Nothing
                    Dim cntrlName = pair(0).Trim()
                    info.ControlsInfo.Add(cntrlName.ToLower(), pair(1).Trim)
                    info.ControlNames.Add(cntrlName)
                ElseIf txt.Substring(pos1, 2) = "'}" Then
                    info.EventHandlers = ParseEventHandlers(txt)
                    Return info
                Else
                    Return Nothing
                End If
            Loop
        End Function

        Public Shared Function ParseEventHandlers(txt As String) As Dictionary(Of String, (ControlName As String, EventName As String))
            '#Events{
            '    Button1: OnMouseLeftDown
            '    Form1: OnMouseLeftUp OnMouseMove
            '    Label1: OnMouseEnter OnKeyUp OnKeyDown
            '    TextBox1: OnMouseLeftUp OnKeyDown
            '}

            Dim events As New Dictionary(Of String, (ControlName As String, EventName As String))

            Dim pos1 = txt.IndexOf("'#Events{")
            If pos1 = -1 Then Return Nothing
            Dim pos2 = txt.IndexOf("'}", pos1)
            If pos2 = -1 Then Return Nothing
            Dim lines = txt.Substring(pos1, pos2 - pos1 + 2).Split({vbCr, vbLf}, StringSplitOptions.RemoveEmptyEntries)
            For i = 1 To lines.Length - 2
                Dim info = lines(i).TrimStart("'"c, " "c).Split(":"c)
                If info Is Nothing OrElse info.Length <> 2 Then Return Nothing
                Dim controlName = info(0)
                Dim eventNames = info(1).Split({" "c}, StringSplitOptions.RemoveEmptyEntries)
                For Each ev In eventNames
                    events(controlName & "_" & ev) = (controlName, ev)
                Next
            Next

            Return events
        End Function

        Public Shared Function GetEventModule(controlName As String, eventName As String) As String
            If EventsInfo(controlName).Contains(eventName) Then
                Return controlName
            Else
                Return NameOf(Control)
            End If
        End Function


    End Class

    Public Class FormInfo
        Public Form As String
        Public ControlsInfo As New Dictionary(Of String, String)
        Public ControlNames As New List(Of String)
        Public EventHandlers As Dictionary(Of String, (ControlName As String, EventName As String))
    End Class


    Public Class ExPropertyAttribute
        Inherits Attribute

    End Class

    Public Class ExMethodAttribute
        Inherits Attribute

    End Class

End Namespace