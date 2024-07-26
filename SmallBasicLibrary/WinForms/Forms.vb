Imports System.Windows.Markup
Imports Microsoft.SmallVisualBasic.Library
Imports Microsoft.SmallVisualBasic.Library.Internal
Imports Wpf = System.Windows.Controls
Imports ControlsDictionay = System.Collections.Generic.Dictionary(Of String, System.Windows.FrameworkElement)
Imports System.Windows.Controls
Imports System.Windows
Imports App = Microsoft.SmallVisualBasic.Library.Internal.SmallBasicApplication
Imports System.Reflection

Namespace WinForms
    ''' <summary>
    ''' This types provides methods to create forms, show them, and get a list of all opened forms in the program.
    ''' </summary>
    <SmallVisualBasicType>
    Public NotInheritable Class Forms
        Friend Shared _forms As New ControlsDictionay
        Friend Shared _controls As New ControlsDictionay

        ''' <summary>
        ''' Used internally by the compiler to stop the calssic timer when the form is closed
        ''' </summary>
        <HideFromIntellisense>
        Public Shared Property TimerParentForm As Primitive

        Public Shared Function GetForm(name As String) As Window
            name = name.ToLower()
            Dim wnd As Window
            Dim dictionary = If(name.Contains("."), _controls, _forms)

            If Not dictionary.ContainsKey(name) Then Return Nothing
            wnd = dictionary(name)

            If wnd Is Nothing Then
                ReportError($"`{name}` is not a form or it is closed", Nothing)
            End If

            Program.ActiveWindow = wnd
            Return wnd
        End Function

        ''' <summary>
        ''' Gets the forms of the current app.
        ''' </summary>
        ''' <param name="loadedFormsOnly">If True, this method will return the loaded forms only, otherwise, it will return all forms defined in this app even if you didn't load them yet.</param>
        ''' <returns>an array containing the names the forms</returns>
        <ReturnValueType(VariableType.Array)>
        Public Shared Function GetForms(loadedFormsOnly As Primitive) As Primitive
            Dim map = New Dictionary(Of Primitive, Primitive)
            Dim num = 1
            Dim formNames As IEnumerable

            If loadedFormsOnly Then
                For Each name In _forms.Keys
                    map(num) = name
                    num += 1
                Next

            ElseIf Not App.IsDebugging Then
                Dim asm = System.Reflection.Assembly.GetEntryAssembly()
                Dim progModule = FormPrefix & "Program"

                For Each frmType In asm.GetTypes()
                    If frmType.Name.StartsWith(FormPrefix) AndAlso frmType.Name <> progModule Then
                        map(num) = frmType.Name.Substring(FormPrefix.Length).ToLower()
                        num += 1
                    End If
                Next

            ElseIf Program.FormNames Is Nothing Then
                Return Nothing
            Else
                For Each name In Program.FormNames
                    map(num) = name.ToLower()
                    num += 1
                Next
            End If

            Return Primitive.ConvertFromMap(map)
        End Function

        <ReturnValueType(VariableType.String)>
        Public Shared Property AppPath As Primitive

        Private Shared _syncLock As New Object

        ''' <summary>
        ''' Loads a form from its xaml file.
        ''' </summary>
        ''' <param name="formName">the name of the form</param>
        ''' <param name="xamlPath">the path pf the xaml file that contains the form design</param>
        ''' <returns>The name of the form</returns>
        <ReturnValueType(VariableType.Form)>
        Public Shared Function LoadForm(formName As Primitive, xamlPath As Primitive) As Primitive
            Dim form_Name = CStr(formName).ToLower()
            If form_Name = "" Then
                Dim ex As New ArgumentException("Form name can't be an empty string.")
                ReportError(ex.Message, ex)
            End If

            If _forms.ContainsKey(form_Name) Then Return form_Name

            SyncLock _syncLock
                App.Invoke(
                    Sub()
                        Try
                            Dim canvas As Canvas
                            Dim xaml = CStr(xamlPath)

                            If xaml = "" Then
                                canvas = LoadContent(form_Name & ".xaml")
                            ElseIf xaml.StartsWith("<") Then
                                xaml = ExpandRelativeImageFiles(xaml, "")
                                Dim stream = New IO.MemoryStream(System.Text.Encoding.UTF8.GetBytes(xaml))
                                canvas = XamlReader.Load(stream)
                            Else
                                canvas = LoadContent(xamlPath)
                            End If

                            canvas.Margin = New Thickness(0)

                            Dim wnd As New Window() With {
                                   .SizeToContent = SizeToContent.WidthAndHeight,
                                   .WindowStartupLocation = WindowStartupLocation.CenterScreen,
                                   .Name = formName,
                                   .Title = Automation.AutomationProperties.GetHelpText(canvas),
                                   .Content = canvas,
                                   .ResizeMode = ResizeMode.CanMinimize,
                                   .Background = canvas.Background,
                                   .MaxWidth = canvas.MaxWidth,
                                   .MaxHeight = canvas.MaxHeight,
                                   .FlowDirection = canvas.FlowDirection
                            }
                            canvas.IsEnabled = True
                            canvas.Visibility = Visibility.Visible
                            AddHandler wnd.Closed, AddressOf Form_Closed
                            AddHandler wnd.Activated, AddressOf Form_Activated
                            _forms(form_Name) = wnd

                            ' Add control names:
                            Dim menu As Menu = Nothing

                            Dim elements = canvas.Children
                            Dim elementNum = -1
                            Do
                                elementNum += 1
                                If elementNum >= elements.Count Then Exit Do

                                Dim ui = elements(elementNum)
                                If TypeOf ui Is Menu Then
                                    menu = ui
                                    Continue Do
                                End If

                                Dim fw = TryCast(ui, FrameworkElement)
                                If fw?.Name = "__FORM__PROPS__" Then
                                    wnd.MinWidth = fw.MinWidth
                                    wnd.MinHeight = fw.MinHeight
                                    wnd.MaxWidth = fw.MaxWidth
                                    wnd.MaxHeight = fw.MaxHeight
                                    wnd.Tag = fw.Tag
                                    wnd.ToolTip = fw.ToolTip
                                    wnd.Left = fw.GetValue(Canvas.LeftProperty)
                                    wnd.Top = fw.GetValue(Canvas.TopProperty)
                                    If Not Double.IsNaN(wnd.Left) OrElse Not Double.IsNaN(wnd.Top) Then
                                        wnd.WindowStartupLocation = WindowStartupLocation.Manual
                                    End If
                                    Continue Do
                                End If

                                Dim lst = TryCast(ui, ItemsControl)
                                If lst IsNot Nothing AndAlso lst.Items.Count = 1 Then
                                    Dim item = lst.Items(0).ToString()
                                    Dim typeNmae = ui.GetType.Name
                                    If item.StartsWith(typeNmae) AndAlso IsNumeric(item.Substring(typeNmae.Length)) Then
                                        lst.Items.Clear()
                                    End If
                                End If

                                Dim controlName = Automation.AutomationProperties.GetName(ui)

                                If controlName = "" Then
                                    controlName = fw.Name
                                    If controlName = "" Then Continue Do
                                End If

                                Dim angle = 0.0
                                Dim rotation As Media.RotateTransform
                                If ui.RenderTransform IsNot Nothing Then
                                    rotation = TryCast(ui.RenderTransform, Media.RotateTransform)
                                    If rotation IsNot Nothing Then angle = rotation.Angle
                                End If

                                Dim key = form_Name & "." & controlName.ToLower()
                                If TypeOf ui Is Wpf.Control Then
                                    _controls(key) = ui
                                    CType(ui, Wpf.Control).Name = controlName
                                    Control.SetAngle(ui, angle)

                                    Dim c = CType(ui, Wpf.Control)
                                    If c.Focusable Then AddHandler c.LostKeyboardFocus, AddressOf Control.OnLostKeyboardFocus
                                    c.SetValue(Control.BorderThicknessProperty, c.BorderThickness)
                                    c.SetValue(Control.BorderBrushProperty, c.BorderBrush)
                                    c.SetValue(Control.TipProperty, c.ToolTip)

                                Else
                                    Dim left = Canvas.GetLeft(ui)
                                    Dim top = Canvas.GetTop(ui)
                                    elements.Remove(ui)
                                    elementNum -= 1

                                    Dim LayoutTransform = fw.LayoutTransform
                                    fw.LayoutTransform = Nothing
                                    ui.RenderTransform = Nothing

                                    Dim lb As New Wpf.Label() With {
                                            .Name = controlName,
                                            .MinWidth = fw.MinWidth,
                                            .MaxWidth = fw.MaxWidth,
                                            .MinHeight = fw.MinHeight,
                                            .MaxHeight = fw.MaxHeight,
                                            .Width = fw.Width,
                                            .Height = fw.Height,
                                            .IsEnabled = fw.IsEnabled,
                                            .Visibility = fw.Visibility,
                                            .FlowDirection = fw.FlowDirection,
                                            .Tag = fw.Tag,
                                            .ToolTip = fw.ToolTip,
                                            .HorizontalContentAlignment = HorizontalAlignment.Stretch,
                                            .VerticalContentAlignment = VerticalAlignment.Stretch,
                                            .Content = ui,
                                            .Background = Nothing,
                                            .LayoutTransform = LayoutTransform,
                                            .RenderTransform = Nothing
                                     }

                                    Panel.SetZIndex(lb, Panel.GetZIndex(fw))
                                    Control.SetAngle(lb, angle)
                                    fw.ClearValue(FrameworkElement.WidthProperty)
                                    fw.ClearValue(FrameworkElement.HeightProperty)
                                    fw.ClearValue(FrameworkElement.MinWidthProperty)
                                    fw.ClearValue(FrameworkElement.MinHeightProperty)
                                    fw.ClearValue(FrameworkElement.MaxWidthProperty)
                                    fw.ClearValue(FrameworkElement.MaxHeightProperty)
                                    fw.ClearValue(FrameworkElement.IsEnabledProperty)
                                    fw.ClearValue(FrameworkElement.VisibilityProperty)
                                    fw.ClearValue(FrameworkElement.FlowDirectionProperty)
                                    fw.ClearValue(FrameworkElement.TagProperty)
                                    fw.ClearValue(FrameworkElement.ToolTipProperty)
                                    elements.Add(lb)
                                    Canvas.SetLeft(lb, left)
                                    Canvas.SetTop(lb, top)
                                    _controls(key) = lb
                                End If

                                SetControlText(ui, key, GetControlText(ui))
                            Loop

                            If menu IsNot Nothing Then
                                elements.Remove(menu)
                                elementNum -= 1
                                Form.AddMenu(wnd, menu)
                                For Each m In menu.Items
                                    AddMenuNames(form_Name, m)
                                Next
                                AddHandler wnd.PreviewKeyDown, AddressOf Form_PreviewKeyDown
                            End If

                        Catch ex As Exception
                            ReportError("LoadForm caused an error: " & ex.Message, ex)
                        End Try
                    End Sub)

            End SyncLock

            ' Ensure Keyboard and Mouse modules are loaded, 
            ' to create a global hanler for the PreviewKeyDown and PreviewMouseWheelEvent events
            Dim __ = Keyboard.LastKey
            __ = Mouse.LastMouseWheelDirection

            Return form_Name
        End Function

        Public Shared Sub EndIfNoForms()
            If _forms.Count = 0 Then Program.End()
        End Sub

        Private Shared Sub Form_Activated(sender As Object, e As EventArgs)
            Program.ActiveWindow = sender
        End Sub

        Friend Shared forceClose As Boolean = False
        Public Const FormPrefix As String = "_SmallVisualBasic_"

        Friend Shared Sub ForceCloseAll()
            forceClose = True
            For Each w As Window In _forms.Values
                Try
                    w.Close()
                Catch
                End Try
            Next

            _forms.Clear()
            _controls.Clear()
            forceClose = False
        End Sub

        Private Shared Sub Form_PreviewKeyDown(sender As Object, e As Input.KeyEventArgs)
            App.Invoke(
                Sub()
                    For Each item In MenuItem.ShortcutHandlers
                        Dim sh = item.Value
                        If sh Is Nothing Then Continue For
                        If e.Key <> sh.key Then Continue For
                        If sh.Ctrl AndAlso (e.KeyboardDevice.Modifiers And Input.ModifierKeys.Control) = 0 Then Continue For
                        If sh.Shift AndAlso (e.KeyboardDevice.Modifiers And Input.ModifierKeys.Shift) = 0 Then Continue For
                        If sh.Alt AndAlso (e.KeyboardDevice.Modifiers And Input.ModifierKeys.Alt) = 0 Then Continue For

                        Dim m = item.Key
                        If m.IsCheckable Then
                            m.IsChecked = Not m.IsChecked
                            ' If there is no OnCkeck handler, call the OnClick handler
                            If sh.Handler IsNot Nothing Then Call sh.Handler()
                        Else
                            Call sh.Handler()
                        End If

                        e.Handled = True
                        Exit Sub
                    Next
                End Sub)
        End Sub

        Private Shared Sub AddMenuNames(form_Name As String, item As Wpf.Control)
            Dim key = form_Name & "." & item.Name.ToLower()
            _controls(key) = item
            If TypeOf item Is Separator Then Return

            Dim m = CType(item, Wpf.MenuItem)
            m.IsCheckable = (LCase(m.Tag) = "true")

            For Each m2 In m.Items
                AddMenuNames(form_Name, m2)
            Next
        End Sub

        Public Shared Function ExpandRelativeImageFiles(xaml As String, fileName As String) As String
            Dim d = If(fileName = "",
                Program.Directory.AsString(),
                IO.Path.GetDirectoryName(fileName)
            ).ToLower().Replace("&", "&amp;") & IO.Path.DirectorySeparatorChar

            xaml = xaml.Replace("Source=""\", $"Source=""{d}")
            xaml = xaml.Replace("FileName=""\", $"FileName=""{d}")
            xaml = xaml.Replace("Source=""/", $"Source=""{d}")
            xaml = xaml.Replace("FileName=""/", $"FileName=""{d}")
            Return xaml
        End Function

        Private Shared Sub SetControlText(control As UIElement, key As String, value As String)
            Try
                CObj(control).Text = value
            Catch
                Try
                    Dim cc = TryCast(control, Wpf.ContentControl)
                    If cc Is Nothing Then Return
                    Dim x = cc.Content
                    If x Is Nothing OrElse TypeOf x Is String Then
                        If TypeOf (control) Is Wpf.Label OrElse TypeOf (control) Is Wpf.Primitives.ButtonBase Then
                            Dim tb = Label.GetTextBlock([Event].GetControlName(control))
                            tb.Text = value
                        Else
                            cc.Content = value
                        End If
                    End If
                Catch
                End Try
            End Try
        End Sub

        Private Shared Function GetControlText(control As UIElement) As String
            Try
                Return CObj(control).Text
            Catch
                Try
                    Dim x = CObj(control).Content
                    If x IsNot Nothing AndAlso TypeOf x Is String Then
                        Return CStr(CObj(control).Content)
                    End If
                Catch
                End Try
            End Try

            Return Automation.AutomationProperties.GetHelpText(control)
        End Function

        ''' <summary>
        ''' Creates a new form with the given name, and adds it to the forms collection.
        ''' </summary>
        ''' <param name="formName">the name of the form</param>
        ''' <returns>the name of the form</returns>
        <ReturnValueType(VariableType.Form)>
        Public Shared Function AddForm(formName As Primitive) As Primitive
            Try
                Dim frm = GetForm(formName)
                If frm Is Nothing Then
                    Dim xaml As String = $"<Canvas Name=""{formName}"" Width=""700"" Height=""500"" xmlns=""http://schemas.microsoft.com/winfx/2006/xaml/presentation""/>"
                    LoadForm(formName, xaml)
                End If

            Catch ex As Exception
                ReportError("AddForm cuased this error: " & ex.Message, ex)
            End Try
            Return formName
        End Function

        Friend Shared Sub Form_Closed(sender As Object, e As EventArgs)
            If forceClose Then Return

            Dim win = CType(sender, Window)
            Dim winName = win.Name.ToLower()
            Dim handler As SmallVisualBasicCallback = Nothing
            Try
                [Event]._senderControl = winName
                Dim keys = Form.ClosedHandlers.Keys
                For i = 0 To Form.ClosedHandlers.Count - 1
                    Dim key = Keys(i)
                    If key.EndsWith(":" & winName) Then
                        handler = Form.ClosedHandlers(key)
                        If handler IsNot Nothing Then
                            Form.ClosedHandlers(key) = Nothing
                            Call handler()
                        End If
                    End If
                Next
                [Event].Handled = False

            Catch ex As Exception
                Helper.ReportError($"The event handler sub `{handler.Method.Name}` fired by the `OnClosed` event, caused this error: {ex.Message}", ex)
            End Try

            App.BeginInvoke(Sub() RemoveFormAndControls(win.Name.ToLower()))
        End Sub

        Friend Shared Sub RemoveFormAndControls(formName As String)
            If _forms.Count = 1 AndAlso _forms.ContainsKey(formName) Then
                ' Prevent errors in debugging mode because of timers after closing the last form
                Program.IsTerminated = True
            End If

            If TimerParentForm.AsString() = formName Then
                Timer.Pause()
                RemoveHandler Timer.Tick, Nothing
            End If

            ' If this control is a WinTimer, remove Tick handlers
            Dim TimerKeys = WinTimer.Timers.Keys
            For i = TimerKeys.Count - 1 To 0 Step -1
                Dim name = TimerKeys(i)
                If name.StartsWith(formName) Then
                    If WinTimer.TickHandlers.ContainsKey(name) Then
                        RemoveHandler WinTimer.Timers(name).Tick, WinTimer.TickHandlers(name)
                        WinTimer.TickHandlers.Remove(name)
                    End If
                    WinTimer.Timers.Remove(name)
                End If
            Next

            _forms.Remove(formName)
            Forms.RemoveEventHandlers(formName)
            formName &= "."

            For i = _controls.Keys.Count - 1 To 0 Step -1
                Dim controlKey = _controls.Keys(i)
                If controlKey.StartsWith(formName) Then
                    Forms.RemoveEventHandlers(controlKey)
                End If
            Next
            If _forms.Count = 0 Then Program.End()
        End Sub

        Private Shared Sub RemoveEventHandlers(controlKey As String)
            Dim RemoveEventHandlers = Control.RemoveEventHandlerActions
            Dim x = $":{controlKey}."
            For j = RemoveEventHandlers.Count - 1 To 0 Step -1
                Dim key = RemoveEventHandlers.Keys(j)
                If key.Contains(x) Then
                    RemoveEventHandlers.Values(j)?.Invoke()
                    RemoveEventHandlers.Remove(key)
                End If
            Next
            _controls.Remove(controlKey)
        End Sub

        Private Shared Function LoadContent(xamlPath As String) As Canvas
            If IO.Path.GetPathRoot(xamlPath) = "" Then
                If AppPath.ToString() <> "" Then
                    xamlPath = IO.Path.Combine(AppPath, xamlPath)
                Else
                    Dim d = Environment.GetCommandLineArgs(0)
                    d = IO.Path.GetDirectoryName(d)
                    Dim xamlPath2 = IO.Path.Combine(d, xamlPath)
                    If IO.File.Exists(xamlPath2) Then xamlPath = xamlPath2
                End If
            End If

            Return GetCanvas(xamlPath)
        End Function

        Private Shared Function GetCanvas(xamlPath As String) As Canvas
            Dim xaml = IO.File.ReadAllText(xamlPath, System.Text.Encoding.UTF8)
            If xaml.Contains("wpf/xaml/WpfDialogs") Then
                xaml = "<Canvas " & "xmlns:mc=""http://schemas.openxmlformats.org/markup-compatibility/2006"" mc:Ignorable=""c""" & xaml.Substring(7)
            End If

            Try
                xaml = ExpandRelativeImageFiles(xaml, xamlPath)
                Dim stream = New IO.MemoryStream(System.Text.Encoding.UTF8.GetBytes(xaml))
                Dim canvas = CType(XamlReader.Load(stream), Canvas)
                Return canvas
            Catch
            End Try

            Return Nothing
        End Function

        ''' <summary>
        ''' Shows a message box dialog.
        ''' Use MsgBox as a shorcut name to show the message box. Ex:
        ''' MsgBox "Hello!"
        ''' </summary>
        ''' <param name="message">the text to dislpay on the message box</param>
        ''' <param name="title">the title to display of the dialog box</param>
        Public Shared Sub ShowMessage(message As Primitive, title As Primitive)
            Try
                SmallBasicApplication.Invoke(
                    Sub()
                        System.Windows.MessageBox.Show(message.ToString(), title.ToString())
                    End Sub)
            Catch ex As Exception
                ReportError("ShowMessage caused this error: " & ex.Message, ex)
            End Try
        End Sub

        Public Shared Event DebugShowForm(formName As String, argsArr As Primitive)

        ''' <summary>
        ''' Shows the form that has the given name if exists in the project.
        ''' </summary>
        ''' <param name="formName">the name of the form.</param>
        ''' <param name="argsArr">any additional data, array, or a dynamic object you want to pass to the form. It will be stored in the ArgsArr property of the form, so you can use it as you want</param>
        ''' <returns>the form name</returns>
        <ReturnValueType(VariableType.Form)>
        Public Shared Function ShowForm(formName As Primitive, argsArr As Primitive) As Primitive
            Dim asm = System.Reflection.Assembly.GetCallingAssembly()

            If App.IsDebugging AndAlso (asm.FullName.StartsWith("sVBCompiler,") OrElse
                     asm.FullName.StartsWith("Unittest,")) Then
                RaiseEvent DebugShowForm(LCase(formName), argsArr)
            Else
                App.Invoke(
                      Sub()
                          Try
                              If Form.GetIsLoaded(formName) Then
                                  DoShowForm(formName, argsArr)
                              Else
                                  Stack.PushValue("_" & formName.AsString().ToLower() & "_argsArr", argsArr)
                                  Form.Initialize(formName, asm)
                              End If
                          Catch ex As Exception
                              Form.ReportSubError(formName, "ShowForm", ex)
                          End Try
                      End Sub)
            End If

            Return formName
        End Function

        Public Shared Sub DoShowForm(formName As Primitive, argsArr As Primitive)
            Form.SetArgsArr(formName, argsArr)
            Form.Show(formName)
            App.Invoke(
                 Sub()
                     Dim wind = GetForm(formName)
                     wind.RaiseEvent(New RoutedEventArgs(Form.OnFormShownEvent))
                 End Sub)
        End Sub

        Public Shared Event DebugShowDialog(formName As String, argsArr As Primitive)

        ''' <summary>
        ''' Loads the form that has the given name if exists in the project, and shows it as a modal dialog.
        ''' </summary>
        ''' <param name="formName">the name of the form.</param>
        ''' <param name="argsArr">any additional data, array, or a dynamic object you want to pass to the form. It will be stored in the Form.ArgsArr property, so you can use it as you want</param>
        ''' <returns>the dialog result that Represents the type of the button that user clicked, like OK, Yes, No, ... etc.</returns>
        <ReturnValueType(VariableType.DialogResult)>
        Public Shared Function ShowDialog(formName As Primitive, argsArr As Primitive) As Primitive
            Dim asm = System.Reflection.Assembly.GetCallingAssembly()

            If App.IsDebugging AndAlso asm.FullName.StartsWith("sVBCompiler,") Then
                RaiseEvent DebugShowDialog(LCase(formName), argsArr)
                Return formName
            End If

            App.Invoke(
                Sub()
                    Try
                        If Form.GetIsLoaded(formName) Then
                            Form.SetArgsArr(formName, argsArr)
                            ShowDialog = Form.ShowDialog(formName)
                        Else
                            Stack.PushValue("_" & formName.AsString().ToLower() & "_argsArr", argsArr)
                            Form.Initialize(formName, asm)
                            Control.SetVisible(formName, False)
                            ShowDialog = Form.ShowDialog(formName)
                        End If
                    Catch ex As Exception
                        Form.ReportSubError(formName, "ShowForm", ex)
                    End Try
                End Sub)
        End Function

    End Class
End Namespace