

Imports Microsoft.SmallBasic.Library

Namespace WinForms
    Public NotInheritable Class PreCompiler

        Private Shared moduleInfo As New Dictionary(Of String, List(Of String))
        Private Shared eventsInfo As New Dictionary(Of String, List(Of String))
        Private Shared deafaultControlEvents As New Dictionary(Of String, String)
        Private Shared types As New List(Of Type)

        Shared Sub New()
            FillModuleMembers(GetType(Forms))
            FillModuleMembers(GetType(Form))
            FillModuleMembers(GetType(Control))
            FillModuleMembers(GetType(TextBox))
            FillModuleMembers(GetType(Label))
            FillModuleMembers(GetType(Button))
            FillModuleMembers(GetType(ListBox))
            FillModuleMembers(GetType(ImageBox))
            FillModuleMembers(GetType(TextEx))
            FillModuleMembers(GetType(MathEx))
            FillModuleMembers(GetType(ArrayEx))
            FillModuleMembers(GetType(ColorEx))

            deafaultControlEvents(NameOf(Form).ToLower()) = "OnShown"
            deafaultControlEvents(NameOf(TextBox).ToLower()) = "OnTextChanged"
            deafaultControlEvents(NameOf(ListBox).ToLower()) = "OnSelection"
        End Sub

        Public Shared Function GetDefaultEvent(controlName As String) As String
            controlName = controlName?.ToLower()
            If controlName <> "" AndAlso deafaultControlEvents.ContainsKey(controlName) Then
                Return deafaultControlEvents(controlName)
            Else
                Return "OnClick"
            End If
        End Function

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

        Public Shared Function GetVarType(variableName As String) As VariableType
            Dim varName = variableName.ToLower()
            Dim varType = VariableType.None

            For Each key In moduleInfo.Keys
                Dim controlName = key.ToLower
                If varName.StartsWith(controlName) OrElse varName.EndsWith(controlName) Then
                    [Enum].TryParse(Of VariableType)(controlName, varType)
                    Return varType
                End If
            Next



            Return varType
        End Function

        Dim PrimativeType As Type = GetType(Primitive)
        Private Const WinFormsNS As String = "Microsoft.SmallBasic.WinForms."
        Private Const LibraryNS As String = "Microsoft.SmallBasic.Library."

        Private Shared Sub FillModuleMembers(t As Type)
            types.Add(t)

            Dim members As New List(Of String)

            For Each m In t.GetMembers()
                members.Add(m.Name.ToLower)
            Next

            moduleInfo(t.Name) = members

            Dim events As New List(Of String)
            For Each e In t.GetEvents()
                If e.EventHandlerType Is GetType(SmallBasicCallback) Then
                    events.Add(e.Name)
                End If
            Next

            eventsInfo(t.Name) = events
        End Sub

        Public Shared Function GetEvents(controlName As String) As List(Of String)
            Dim events As New List(Of String)
            If controlName = NameOf(Forms) Then Return events

            If controlName <> NameOf(ImageBox) Then
                For Each e In eventsInfo(NameOf(Control))
                    events.Add(e)
                Next
            End If

            If controlName <> NameOf(Control) Then
                For Each e In eventsInfo(controlName)
                    If Not events.Contains(e) Then events.Add(e)
                Next
            End If

            events.Sort()
            Return events
        End Function

        Public Shared Function GetMethodInfo(
                         controlName As String,
                         varType As VariableType,
                         methodName As String
                   ) As ([Module] As String, ParamsCount As Integer)

            If controlName = "" Then
                Select Case varType
                    Case VariableType.String
                        controlName = NameOf(TextEx)
                    Case VariableType.Double
                        controlName = NameOf(MathEx)
                    Case VariableType.Array
                        controlName = NameOf(ArrayEx)
                    Case VariableType.Color
                        controlName = NameOf(ColorEx)
                    Case Else
                        Return ("", 0)
                End Select
            End If

            Dim method = methodName.ToLower()
            Dim moduleName = ""

            If moduleInfo(controlName).Contains(method) Then
                moduleName = controlName

            ElseIf controlName = NameOf(ImageBox) Then
                Return ("", 0)

            ElseIf moduleInfo(NameOf(Control)).Contains(method) Then
                moduleName = NameOf(Control)

            ElseIf moduleInfo(NameOf(Forms)).Contains(method) Then
                moduleName = NameOf(Forms)

            Else
                Return ("", 0)
            End If

            Dim t = Type.GetType(WinFormsNS & moduleName)
            Dim params = t?.GetMethods.Where(
                   Function(m) m.IsPublic AndAlso m.Name.ToLower() = method
            ).FirstOrDefault?.GetParameters()

            Return (moduleName, If(params Is Nothing, 0, params.Length))
        End Function

        Public Shared Function ParseFormHints(txt As String) As FormInfo
            '@Form Hints:
            '#Form1{
            '    TextBox1: TextBox
            '    btnOk: Button
            '    lblError: Label
            '}

            Dim pos1 = txt.IndexOf("'@Form Hints:")
            If pos1 = -1 Then Return Nothing

            Dim lines = txt.Substring(pos1).Split({vbCr, vbLf}, StringSplitOptions.RemoveEmptyEntries)
            If lines.Count < 3 Then Return Nothing


            If Not (lines(1).StartsWith("'#") AndAlso lines(1).EndsWith("{")) Then Return Nothing

            Dim info As New FormInfo
            info.Form = lines(1).Substring(2, lines(1).Length - 3).Trim()
            If info.Form = "" Then Return Nothing
            info.ControlsInfo(info.Form.ToLower()) = "Form"
            info.ControlNames.Add(info.Form)

            For i = 2 To lines.Count - 1
                If lines(i) = "'}" Then
                    info.EventHandlers = ParseEventHandlers(txt)
                    Return info
                End If

                If Not lines(i).StartsWith("'    ") Then Return Nothing

                Dim hint = lines(i).Substring(5)
                Dim pair = hint.Split(":"c)
                If pair Is Nothing OrElse pair.Length <> 2 Then Return Nothing

                Dim cntrlName = pair(0).Trim()
                Dim c = cntrlName.ToLower()
                If Not info.ControlsInfo.ContainsKey(c) Then
                    info.ControlsInfo.Add(c, pair(1).Trim)
                    info.ControlNames.Add(cntrlName)
                End If

            Next
            Return Nothing
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
            If eventsInfo(controlName).Contains(eventName) Then
                Return controlName
            Else
                Return NameOf(Control)
            End If
        End Function

        Public Shared Function GetModuleName(name As String) As String
            If moduleInfo.ContainsKey(name) Then Return name
            Return NameOf(Control)
        End Function

        Public Shared Function GetModuleFromVarName(varName As String) As String
            varName = varName.ToLower

            For Each key In moduleInfo.Keys
                Dim controlName = key.ToLower
                If varName.StartsWith(controlName) OrElse varName.EndsWith(controlName) Then
                    Return key
                End If
            Next
            Return ""
        End Function

        Public Shared Function GetTypes() As List(Of Type)
            Return types
        End Function


    End Class

    Public Class FormInfo
        Public Form As String
        Public ControlsInfo As New Dictionary(Of String, String)
        Public ControlNames As New List(Of String)
        Public EventHandlers As Dictionary(Of String, (ControlName As String, EventName As String))
    End Class


End Namespace