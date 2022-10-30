

Imports Microsoft.SmallVisualBasic.Library

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
            FillModuleMembers(GetType(ListBox))
            FillModuleMembers(GetType(ComboBox))
            FillModuleMembers(GetType(CheckBox))
            FillModuleMembers(GetType(RadioButton))
            FillModuleMembers(GetType(ToggleButton))
            FillModuleMembers(GetType(Button))
            FillModuleMembers(GetType(MenuItem))
            FillModuleMembers(GetType(MainMenu))
            FillModuleMembers(GetType(DatePicker))
            FillModuleMembers(GetType(ProgressBar))
            FillModuleMembers(GetType(Slider))
            FillModuleMembers(GetType(ScrollBar))
            FillModuleMembers(GetType(ImageBox))
            FillModuleMembers(GetType(WinTimer))
            FillModuleMembers(GetType(TextEx))
            FillModuleMembers(GetType(MathEx))
            FillModuleMembers(GetType(ArrayEx))
            FillModuleMembers(GetType(ColorEx))
            FillModuleMembers(GetType(DateEx))

            deafaultControlEvents(NameOf(Form).ToLower()) = "OnShown"
            deafaultControlEvents(NameOf(TextBox).ToLower()) = "OnTextChanged"
            deafaultControlEvents(NameOf(ListBox).ToLower()) = "OnSelection"
            deafaultControlEvents(NameOf(ComboBox).ToLower()) = "OnSelection"
            deafaultControlEvents(NameOf(CheckBox).ToLower()) = "OnCheck"
            deafaultControlEvents(NameOf(RadioButton).ToLower()) = "OnCheck"
            deafaultControlEvents(NameOf(ToggleButton).ToLower()) = "OnCheck"
            deafaultControlEvents(NameOf(Slider).ToLower()) = "OnSlide"
            deafaultControlEvents(NameOf(ScrollBar).ToLower()) = "OnScroll"
            deafaultControlEvents(NameOf(DatePicker).ToLower()) = "OnSelection"
        End Sub

        Public Shared Function GetTypeName(varType As VariableType) As String
            Select Case varType

                Case VariableType.String
                    Return NameOf(TextEx)

                Case VariableType.Double
                    Return NameOf(MathEx)

                Case VariableType.Array
                    Return NameOf(ArrayEx)

                Case VariableType.Color
                    Return NameOf(ColorEx)

                Case VariableType.Date
                    Return NameOf(DateEx)

                Case VariableType.Any, VariableType.Boolean,
                         VariableType.Key, VariableType.DialogResult,
                         VariableType.ControlType
                    Return ""

                Case Else
                    Return varType.ToString()

            End Select

        End Function

        Public Shared Function GetTypeDisplayName(type As Type) As String
            If type Is Nothing Then Return ""

            Select Case type.Name
                Case NameOf(ArrayEx)
                    Return NameOf(Array)

                Case NameOf(MathEx)
                    Return NameOf(Math)

                Case NameOf(TextEx)
                    Return NameOf(Text)

                Case NameOf(ColorEx)
                    Return NameOf(Color)

                Case NameOf(DateEx)
                    Return NameOf([Date])
                Case Else
                    Return type.Name
            End Select
        End Function

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


        Private Shared typeShortcuts() As ShortcutInfo = {
                New ShortcutInfo("Str", VariableType.String),
                New ShortcutInfo("Arr", VariableType.Array),
                New ShortcutInfo("Dbl", VariableType.Double),
                New ShortcutInfo("Color", VariableType.Color),
                New ShortcutInfo("Key", VariableType.Key),
                New ShortcutInfo("Dialog", VariableType.DialogResult),
                New ShortcutInfo("TypeName", VariableType.ControlType),
                New ShortcutInfo("ControlType", VariableType.ControlType),
                New ShortcutInfo("Date", VariableType.Date),
                New ShortcutInfo("Control", VariableType.Control),
                New ShortcutInfo("Form", VariableType.Form),
                New ShortcutInfo("TextBox", VariableType.TextBox),
                New ShortcutInfo("Label", VariableType.Label),
                New ShortcutInfo("ListBox", VariableType.ListBox),
                New ShortcutInfo("ComboBox", VariableType.ComboBox),
                New ShortcutInfo("CheckBox", VariableType.CheckBox),
                New ShortcutInfo("RadioButton", VariableType.RadioButton),
                New ShortcutInfo("ToggleButton", VariableType.ToggleButton),
                New ShortcutInfo("Button", VariableType.Button),
                New ShortcutInfo("MenuItem", VariableType.MenuItem),
                New ShortcutInfo("MainMenu", VariableType.MainMenu),
                New ShortcutInfo("DatePicker", VariableType.DatePicker),
                New ShortcutInfo("ProgressBar", VariableType.ProgressBar),
                New ShortcutInfo("Slider", VariableType.Slider),
                New ShortcutInfo("ScrollBar", VariableType.ScrollBar),
                New ShortcutInfo("ImageBox", VariableType.ImageBox),
                New ShortcutInfo("Timer", VariableType.WinTimer)
    }

        Public Shared Function GetVarType(variableName As String) As VariableType
            variableName = variableName.Trim("_")
            Dim varName = variableName.ToLower()

            For Each sh In typeShortcuts
                Dim lowcaseShotrcut = sh.Shortcut(0).ToString.ToLower() & sh.Shortcut.Substring(1)
                Dim n = sh.Shortcut.Length

                If variableName.StartsWith(sh.Shortcut) Then Return sh.Type
                If variableName.EndsWith(sh.Shortcut) Then Return sh.Type

                If variableName.StartsWith(lowcaseShotrcut) Then
                    If variableName.Length = n Then Return sh.Type

                    Dim x = variableName(n)
                    If x = "_" OrElse IsNumeric(x) OrElse Char.IsUpper(x) Then
                        Return sh.Type
                    End If
                End If

                If variableName.EndsWith(lowcaseShotrcut) Then
                    Dim x = variableName(variableName.Length - 1 - n)
                    If x = "_" OrElse IsNumeric(x) Then Return sh.Type
                End If
            Next

            Return VariableType.Any
        End Function

        Dim PrimativeType As Type = GetType(Primitive)
        Private Const WinFormsNS As String = "Microsoft.SmallVisualBasic.WinForms."
        Private Const LibraryNS As String = "Microsoft.SmallVisualBasic.Library."

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

        Public Shared Function FillControlDefaultProperties() As Dictionary(Of String, String)
            Dim defaultProperties As New Dictionary(Of String, String)
            defaultProperties(NameOf(Form)) = "text"
            defaultProperties(NameOf(Control)) = "name"
            defaultProperties(NameOf(TextBox)) = "text"
            defaultProperties(NameOf(Label)) = "text"
            defaultProperties(NameOf(ListBox)) = "additem"
            defaultProperties(NameOf(ComboBox)) = "additem"
            defaultProperties(NameOf(CheckBox)) = "checked"
            defaultProperties(NameOf(RadioButton)) = "checked"
            defaultProperties(NameOf(ToggleButton)) = "checked"
            defaultProperties(NameOf(Button)) = "enabled"
            defaultProperties(NameOf(MenuItem)) = "additem"
            defaultProperties(NameOf(MainMenu)) = "additem"
            defaultProperties(NameOf(DatePicker)) = "selecteddate"
            defaultProperties(NameOf(ProgressBar)) = "maximum"
            defaultProperties(NameOf(Slider)) = "maximum"
            defaultProperties(NameOf(ScrollBar)) = "maximum"
            defaultProperties(NameOf(ImageBox)) = "filename"
            defaultProperties(NameOf(WinTimer)) = "interval"
            Return defaultProperties
        End Function

        Public Shared Function GetEvents(controlName As String) As List(Of String)
            Dim events As New List(Of String)
            If controlName = NameOf(Forms) Then Return events

            If controlName <> NameOf(ImageBox) AndAlso controlName <> NameOf(WinTimer) Then
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

        Public Shared Function ContainsEvent(controlName As String, eventName As String) As Boolean
            If controlName = NameOf(Forms) Then Return False

            eventName = eventName.ToLower()

            If controlName <> NameOf(ImageBox) AndAlso controlName <> NameOf(WinTimer) Then
                For Each e In eventsInfo(NameOf(Control))
                    If e.ToLower = eventName Then Return True
                Next
            End If

            If eventsInfo.ContainsKey(controlName) Then
                For Each e In eventsInfo(controlName)
                    If e.ToLower() = eventName Then Return True
                Next
            End If

            Return False
        End Function

        Public Shared Function GetMethodInfo(
                         controlName As String,
                         varType As VariableType,
                         methodName As String
                   ) As MethodInformation

            If controlName = "" Then
                controlName = GetTypeName(varType)
                If controlName = "" Then Return New MethodInformation("", 0)
            End If

            Dim method = methodName.ToLower()
            Dim moduleName = ""

            If moduleInfo(controlName).Contains(method) Then
                moduleName = controlName

            ElseIf varType < VariableType.Control AndAlso varType <> VariableType.Any Then
                Return New MethodInformation("", 0)

            ElseIf controlName = NameOf(ImageBox) OrElse controlName = NameOf(WinTimer) Then
                Return New MethodInformation("", 0)

            ElseIf moduleInfo(NameOf(Control)).Contains(method) Then
                moduleName = NameOf(Control)

            ElseIf moduleInfo(NameOf(Forms)).Contains(method) Then
                moduleName = NameOf(Forms)

            Else
                Return New MethodInformation("", 0)
            End If

            Dim t = Type.GetType(WinFormsNS & moduleName)
            Dim params = t?.GetMethods.Where(
                   Function(m) m.IsPublic AndAlso m.Name.ToLower() = method
            ).FirstOrDefault?.GetParameters()

            Return New MethodInformation(moduleName, If(params Is Nothing, 0, params.Length))
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
            info.ControlsInfo("me") = "Form"
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

        Public Shared Function ParseEventHandlers(txt As String) As Dictionary(Of String, EventInformation)
            '#Events{
            '    Button1: OnMouseLeftDown
            '    Form1: OnMouseLeftUp OnMouseMove
            '    Label1: OnMouseEnter OnKeyUp OnKeyDown
            '    TextBox1: OnMouseLeftUp OnKeyDown
            '}

            Dim events As New Dictionary(Of String, EventInformation)

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
                    events(controlName & "_" & ev) = New EventInformation(controlName, ev)
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
            Dim controls = moduleInfo.Keys

            ' start from 1 to skip Forms module
            For i = 1 To controls.Count - 1
                Dim controlName = controls(i).ToLower
                If varName.StartsWith(controlName) OrElse varName.EndsWith(controlName) Then
                    Return controls(i)
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
        Public EventHandlers As Dictionary(Of String, EventInformation)
    End Class

    Friend Structure ShortcutInfo
        Public Shortcut As String
        Public Type As VariableType

        Public Sub New(shortcut As String, type As VariableType)
            Me.Shortcut = shortcut
            Me.Type = type
        End Sub

    End Structure

    Public Structure MethodInformation
        Public [Module] As String
        Public ParamsCount As Integer

        Public Sub New([module] As String, paramsCount As Integer)
            Me.Module = [module]
            Me.ParamsCount = paramsCount
        End Sub
    End Structure

    Public Structure EventInformation
        Public ControlName As String
        Public EventName As String

        Public Sub New(controlName As String, eventName As String)
            Me.ControlName = controlName
            Me.EventName = eventName
        End Sub
    End Structure
End Namespace