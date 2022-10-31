Imports Microsoft.SmallVisualBasic.Library
Imports Wpf = System.Windows.Controls
Imports App = Microsoft.SmallVisualBasic.Library.Internal.SmallBasicApplication
Imports System.Windows.Media
Imports System.Windows

Namespace WinForms

    <SmallBasicType>
    <HideFromIntellisense>
    Public NotInheritable Class Form

        Shared Sub ReportSubError(formName As String, memberName As String, ex As Exception)
            Helper.ReportError($"Calling {formName}.{memberName} caused an error: {vbCrLf}{ex.Message}", ex)
        End Sub

        Shared Sub ReportError(formName As String, memberName As String, ex As Exception)
            Helper.ReportError($"Reading {formName}.{memberName} caused an error: {vbCrLf}{ex.Message}", ex)
        End Sub

        Shared Sub ReportError(formName As String, memberName As String, value As String, ex As Exception)
            Helper.ReportError($"Sending `{value}` to {formName}.{memberName} caused an error: {vbCrLf}{ex.Message}", ex)
        End Sub

        Public Shared Sub Initialize(formName As String, asm As System.Reflection.Assembly)
            App.Invoke(
                Sub()
                    Try
                        Dim frmName = "_SmallVisualBasic_" & formName.ToLower()
                        Dim frm = asm.GetType(frmName)
                        Dim initializeMethod = frm?.GetMethod("Initialize", System.Reflection.BindingFlags.Public Or System.Reflection.BindingFlags.Static)
                        initializeMethod?.Invoke(Nothing, Nothing)

                    Catch ex As Exception
                        ReportSubError(formName, "Initialize", ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Adds a new TextBox control to the form
        ''' </summary>
        ''' <param name="textBoxName">A unigue name of the new TextBox.</param>
        ''' <param name="left">The X-pos of the control.</param>
        ''' <param name="top">The Y-pos of the control.</param>
        ''' <param name="width">The width of the control.</param>
        ''' <param name="height">The height of the control.</param>
        ''' <returns>The neame of the textBox</returns>
        <ReturnValueType(VariableType.TextBox)>
        <ExMethod>
        Public Shared Function AddTexBox(
                         formName As Primitive,
                         textBoxName As Primitive,
                         left As Primitive, top As Primitive,
                         width As Primitive, height As Primitive) As Primitive

            Dim key = ValidateArgs(formName, textBoxName)
            App.Invoke(
                 Sub()
                     Try
                         Dim frm = CType(Forms._forms(CStr(formName).ToLower), Window)

                         Dim textBox1 As New Wpf.TextBox With {
                           .Name = textBoxName,
                           .Width = width,
                           .Height = height,
                           .VerticalScrollBarVisibility = Wpf.ScrollBarVisibility.Auto,
                           .HorizontalScrollBarVisibility = Wpf.ScrollBarVisibility.Auto
                         }

                         Wpf.Canvas.SetLeft(textBox1, left)
                         Wpf.Canvas.SetTop(textBox1, top)

                         Dim cnv = GetCanvas(frm)
                         cnv.Children.Add(textBox1)
                         Forms._controls(key) = textBox1

                     Catch ex As Exception
                         ReportSubError(formName, "AddTexBox", ex)
                     End Try
                 End Sub)

            Return key
        End Function

        ''' <summary>
        ''' Adds a new Label control to the form
        ''' </summary>
        ''' <param name="labelName">A unigue name of the new Label.</param>
        ''' <param name="left">The X-pos of the control.</param>
        ''' <param name="top">The Y-pos of the control.</param>
        ''' <param name="width">The width of the control.</param>
        ''' <param name="height">The height of the control.</param>
        ''' <returns>The neame of the label</returns>
        <ReturnValueType(VariableType.Label)>
        <ExMethod>
        Public Shared Function AddLabel(formName As Primitive,
                         labelName As Primitive,
                         left As Primitive, top As Primitive,
                         width As Primitive, height As Primitive) As Primitive

            Dim key = ValidateArgs(formName, labelName)
            App.Invoke(
                    Sub()
                        Try
                            Dim frm = CType(Forms._forms(CStr(formName).ToLower), System.Windows.Window)

                            Dim label1 As New Wpf.Label With {
                               .Name = labelName,
                               .Width = width,
                               .Height = height
                            }

                            Wpf.Canvas.SetLeft(label1, left)
                            Wpf.Canvas.SetTop(label1, top)

                            Dim cnv = GetCanvas(frm)
                            cnv.Children.Add(label1)
                            Forms._controls(key) = label1

                        Catch ex As Exception
                            ReportSubError(formName, "AddLabel", ex)
                        End Try
                    End Sub)

            Return key
        End Function

        ''' <summary>
        ''' Adds a new IamgeBox control to the form
        ''' </summary>
        ''' <param name="imageBoxName">A unigue name of the new ImageBox.</param>
        ''' <param name="left">The X-pos of the control.</param>
        ''' <param name="top">The Y-pos of the control.</param>
        ''' <param name="width">The width of the control.</param>
        ''' <param name="height">The height of the control.</param>
        ''' <param name="fileName">the path of the image file</param>
        ''' <returns>The neame of the ImageBox</returns>
        <ReturnValueType(VariableType.ImageBox)>
        <ExMethod>
        Public Shared Function AddImageBox(
                         formName As Primitive,
                         imageBoxName As Primitive,
                         left As Primitive,
                         top As Primitive,
                         width As Primitive,
                         height As Primitive,
                         fileName As Primitive
                   ) As Primitive

            Dim key = ValidateArgs(formName, imageBoxName)
            App.Invoke(
                    Sub()
                        Try
                            Dim frm = CType(Forms._forms(CStr(formName).ToLower), System.Windows.Window)
                            Dim img As New Wpf.Image With {
                                .Name = imageBoxName,
                                .Width = width,
                                .Height = height
                            }

                            If Not IO.Path.IsPathRooted(fileName) Then
                                Dim path = Environment.GetCommandLineArgs()(0)
                                path = IO.Path.GetDirectoryName(path)
                                fileName = IO.Path.Combine(path, fileName)
                            End If
                            img.Source = New Imaging.BitmapImage(New Uri(fileName))

                            Wpf.Canvas.SetLeft(img, left)
                            Wpf.Canvas.SetTop(img, top)

                            Dim cnv = GetCanvas(frm)
                            cnv.Children.Add(img)
                            Forms._controls(key) = img

                        Catch ex As Exception
                            ReportSubError(formName, "AddImageBox", ex)
                        End Try
                    End Sub)

            Return key
        End Function

        ''' <summary>
        ''' Adds a new Timer control to the form
        ''' </summary>
        ''' <param name="timerName">A unigue name of the new Timer.</param>
        ''' <param name="interval">The delay time in milliseconds between ticks</param>
        ''' <returns>The neame of the Timer</returns>
        <ReturnValueType(VariableType.WinTimer)>
        <ExMethod>
        Public Shared Function AddTimer(
                         formName As Primitive,
                         timerName As Primitive,
                         interval As Primitive
                   ) As Primitive

            Dim key = ValidateArgs(formName, timerName)
            App.Invoke(
                    Sub()
                        Try
                            Dim frm = CType(Forms._forms(CStr(formName).ToLower), Window)
                            Dim t As New Threading.DispatcherTimer() With {
                                .Tag = timerName,
                                .Interval = TimeSpan.FromMilliseconds(interval)
                            }

                            WinTimer.Timers(key) = t
                            t.Start()

                        Catch ex As Exception
                            ReportSubError(formName, "AddTimer", ex)
                        End Try
                    End Sub)

            Return key
        End Function


        Friend Shared Function ValidateArgs(formName As String, controlName As String) As String
            If formName = "" Then
                Dim msg = "`formName` can't be empty string"
                Dim ex As New ArgumentException(msg)
                Helper.ReportError(msg, ex)
                Throw ex
            End If

            If controlName = "" Then
                Dim msg = "Control name can't be empty string"
                Dim ex As New ArgumentException(msg)
                Helper.ReportError(msg, ex)
                Throw ex
            End If

            Dim frmName = formName.ToString().ToLower
            If Not Forms._forms.ContainsKey(frmName) Then
                Dim msg = $"The Form `{formName}`doesn't exist"
                Dim ex As New ArgumentException(msg)
                Helper.ReportError(msg, ex)
                Throw ex
            End If

            Dim key = frmName & "." & controlName.ToString().ToLower()
            If Forms._controls.ContainsKey(key) Then
                Dim msg = $"There is another control with the name '{controlName}' on the ''{formName} form"
                Dim ex As New ArgumentException(msg)
                Helper.ReportError(msg, ex)
                Throw ex
            End If

            Return key
        End Function

        ''' <summary>
        ''' Adds a main menu to the current form. If there is already a main menu, it will be replaced.
        ''' </summary>
        ''' <param name="menuName">The name of the main menu</param>
        ''' <returns>The menu item that have been added.</returns>

        <ExMethod>
        <ReturnValueType(VariableType.MainMenu)>
        Public Shared Function AddMainMenu(
                         formName As Primitive,
                         menuName As Primitive
                    ) As Primitive

            Dim key = ValidateArgs(formName, menuName)

            App.Invoke(
                Sub()
                    Try
                        Dim frm = CType(Forms._forms(CStr(formName).ToLower), System.Windows.Window)
                        Dim memu As New Wpf.Menu()
                        memu.Name = menuName
                        Forms._controls(key) = memu

                        If TypeOf frm.Content Is Wpf.Canvas Then
                            Dim canvas = frm.Content
                            Dim sp As New Wpf.StackPanel()
                            sp.Orientation = Wpf.Orientation.Vertical
                            frm.Content = sp
                            sp.Children.Add(memu)
                            sp.Children.Add(canvas)
                        Else
                            Dim sp = CType(frm.Content, Wpf.StackPanel)
                            sp.Children(0) = memu
                        End If

                    Catch ex As Exception
                        Control.ReportSubError(menuName, "AddItem", ex)
                    End Try
                End Sub)

            Return key
        End Function


        ''' <summary>
        ''' Adds a new Button control to the form
        ''' </summary>
        ''' <param name="buttonName">A unigue name of the new Button.</param>
        ''' <param name="left">The X-pos of the control.</param>
        ''' <param name="top">The Y-pos of the control.</param>
        ''' <param name="width">The width of the control.</param>
        ''' <param name="height">The height of the control.</param>
        ''' <returns>The neame of the button</returns>
        <ReturnValueType(VariableType.Button)>
        <ExMethod>
        Public Shared Function AddButton(
                         formName As Primitive,
                         buttonName As Primitive,
                         left As Primitive,
                         top As Primitive,
                         width As Primitive,
                         height As Primitive
                    ) As Primitive

            Dim key = ValidateArgs(formName, buttonName)
            App.Invoke(
                  Sub()
                      Try
                          Dim frm = CType(Forms._forms(CStr(formName).ToLower), System.Windows.Window)

                          Dim button1 As New Wpf.Button With {
                               .Name = buttonName,
                               .Width = width,
                               .Height = height
                          }

                          Wpf.Canvas.SetLeft(button1, left)
                          Wpf.Canvas.SetTop(button1, top)

                          Dim cnv = GetCanvas(frm)
                          cnv.Children.Add(button1)
                          Forms._controls(key) = button1

                      Catch ex As Exception
                          ReportSubError(formName, "AddButton", ex)
                      End Try
                  End Sub)

            Return key
        End Function

        ''' <summary>
        ''' Adds a new ToggleButton control to the form
        ''' </summary>
        ''' <param name="toggleButtonName">A unigue name of the new ToggleButton.</param>
        ''' <param name="left">The X-pos of the control.</param>
        ''' <param name="top">The Y-pos of the control.</param>
        ''' <param name="width">The width of the control.</param>
        ''' <param name="height">The height of the control.</param>
        ''' <returns>The neame of the toggleButton</returns>
        <ReturnValueType(VariableType.ToggleButton)>
        <ExMethod>
        Public Shared Function AddToggleButton(
                         formName As Primitive,
                         toggleButtonName As Primitive,
                         left As Primitive,
                         top As Primitive,
                         width As Primitive,
                         height As Primitive
                    ) As Primitive

            Dim key = ValidateArgs(formName, toggleButtonName)
            App.Invoke(
                  Sub()
                      Try
                          Dim frm = CType(Forms._forms(CStr(formName).ToLower), System.Windows.Window)

                          Dim toggleButton1 As New Wpf.Primitives.ToggleButton With {
                               .Name = toggleButtonName,
                               .Width = width,
                               .Height = height
                          }

                          Wpf.Canvas.SetLeft(toggleButton1, left)
                          Wpf.Canvas.SetTop(toggleButton1, top)

                          Dim cnv = GetCanvas(frm)
                          cnv.Children.Add(toggleButton1)
                          Forms._controls(key) = toggleButton1

                      Catch ex As Exception
                          ReportSubError(formName, "AddToggleButton", ex)
                      End Try
                  End Sub)

            Return key
        End Function


        ''' <summary>
        ''' Adds a new CheckBox control to the form
        ''' </summary>
        ''' <param name="checkBoxName">A unigue name of the new CheckBox.</param>
        ''' <param name="left">The X-pos of the control.</param>
        ''' <param name="top">The Y-pos of the control.</param>
        ''' <param name="text">the text to desply on the CheckBox</param>
        ''' <param name="checked">The value to set to the Checked property</param>
        ''' <returns>The neame of the button</returns>
        <ReturnValueType(VariableType.CheckBox)>
        <ExMethod>
        Public Shared Function AddCheckBox(
                         formName As Primitive,
                         checkBoxName As Primitive,
                         left As Primitive,
                         top As Primitive,
                         text As Primitive,
                         checked As Primitive
                   ) As Primitive

            Dim key = ValidateArgs(formName, checkBoxName)
            App.Invoke(
                  Sub()
                      Try
                          Dim frm = CType(Forms._forms(CStr(formName).ToLower), System.Windows.Window)

                          Dim ch As New Wpf.CheckBox With {
                               .Name = checkBoxName,
                               .Content = text
                          }

                          ' Don't use if() expression, because it will return a new primitive not Nothing
                          If checked.AsString() = "" Then
                              ch.IsChecked = Nothing
                          Else
                              ch.IsChecked = checked
                          End If

                          Wpf.Canvas.SetLeft(ch, left)
                          Wpf.Canvas.SetTop(ch, top)

                          Dim cnv = GetCanvas(frm)
                          cnv.Children.Add(ch)
                          Forms._controls(key) = ch

                      Catch ex As Exception
                          ReportSubError(formName, "AddCheckBox", ex)
                      End Try
                  End Sub)

            Return key
        End Function

        ''' <summary>
        ''' Adds a new RadioButton control to the form
        ''' </summary>
        ''' <param name="radioButtonName">A unigue name of the new RadioButton.</param>
        ''' <param name="left">The X-pos of the control.</param>
        ''' <param name="top">The Y-pos of the control.</param>
        ''' <param name="text">The text to desply on the RadioButton</param>
        ''' <param name="groupName">The name of the group to add the button to</param>
        ''' <param name="checked">The value to set to the Checked property</param>
        ''' <returns>The neame of the button</returns>
        <ReturnValueType(VariableType.RadioButton)>
        <ExMethod>
        Public Shared Function AddRadioButton(
                         formName As Primitive,
                         radioButtonName As Primitive,
                         left As Primitive,
                         top As Primitive,
                         text As Primitive,
                         groupName As Primitive,
                         checked As Primitive
                   ) As Primitive

            Dim key = ValidateArgs(formName, radioButtonName)
            App.Invoke(
                  Sub()
                      Try
                          Dim frm = CType(Forms._forms(CStr(formName).ToLower), System.Windows.Window)

                          Dim rd As New Wpf.RadioButton With {
                               .Name = radioButtonName,
                               .Content = text,
                               .GroupName = groupName,
                               .IsChecked = checked
                          }

                          Wpf.Canvas.SetLeft(rd, left)
                          Wpf.Canvas.SetTop(rd, top)

                          Dim cnv = GetCanvas(frm)
                          cnv.Children.Add(rd)
                          Forms._controls(key) = rd

                      Catch ex As Exception
                          ReportSubError(formName, "AddRadioButton", ex)
                      End Try
                  End Sub)

            Return key
        End Function


        ''' <summary>
        ''' Adds a new ListBox control to the form
        ''' </summary>
        ''' <param name="listBoxName">A unigue name of the new ListBox.</param>
        ''' <param name="left">The X-pos of the control.</param>
        ''' <param name="top">The Y-pos of the control.</param>
        ''' <param name="width">The width of the control.</param>
        ''' <param name="height">The height of the control.</param>
        ''' <returns>The neame of the listBox</returns>
        <ReturnValueType(VariableType.ListBox)>
        <ExMethod>
        Public Shared Function AddListBox(formName As Primitive,
                         listBoxName As Primitive,
                         left As Primitive,
                         top As Primitive,
                         width As Primitive,
                         height As Primitive
                  ) As Primitive

            Dim key = ValidateArgs(formName, listBoxName)
            App.Invoke(
                  Sub()
                      Try
                          Dim frm = CType(Forms._forms(CStr(formName).ToLower), System.Windows.Window)

                          Dim listBox1 As New Wpf.ListBox With {
                                 .Name = listBoxName,
                                 .Width = width,
                                 .Height = height
                          }

                          Wpf.Canvas.SetLeft(listBox1, left)
                          Wpf.Canvas.SetTop(listBox1, top)

                          Dim cnv = GetCanvas(frm)
                          cnv.Children.Add(listBox1)
                          Forms._controls(key) = listBox1

                      Catch ex As Exception
                          ReportSubError(formName, "AddListBox", ex)
                      End Try
                  End Sub)

            Return key
        End Function

        ''' <summary>
        ''' Adds a new DatePicker control to the form
        ''' </summary>
        ''' <param name="datePickerName">A unigue name of the new DatePicker.</param>
        ''' <param name="left">The X-pos of the control.</param>
        ''' <param name="top">The Y-pos of the control.</param>
        ''' <param name="width">The width of the control.</param>
        ''' <param name="selectedDate">the date that will be selected in the control</param>
        ''' <returns>The neame of the DatePicker</returns>
        <ReturnValueType(VariableType.DatePicker)>
        <ExMethod>
        Public Shared Function AddDatePiker(
                         formName As Primitive,
                         datePickerName As Primitive,
                         left As Primitive,
                         top As Primitive,
                         width As Primitive,
                         selectedDate As Primitive
                    ) As Primitive

            Dim key = ValidateArgs(formName, datePickerName)
            App.Invoke(
                  Sub()
                      Try
                          Dim frm = CType(Forms._forms(CStr(formName).ToLower), System.Windows.Window)

                          Dim dp As New Wpf.DatePicker With {
                               .Name = datePickerName,
                               .Width = width,
                               .SelectedDate = selectedDate
                          }

                          Wpf.Canvas.SetLeft(dp, left)
                          Wpf.Canvas.SetTop(dp, top)

                          Dim cnv = GetCanvas(frm)
                          cnv.Children.Add(dp)
                          Forms._controls(key) = dp

                      Catch ex As Exception
                          ReportSubError(formName, "AddDatePiker", ex)
                      End Try
                  End Sub)

            Return key
        End Function

        ''' <summary>
        ''' Adds a new ProgressBar control to the form
        ''' </summary>
        ''' <param name="progressBarName">A unigue name of the new ProgressBar.</param>
        ''' <param name="left">The X-pos of the control.</param>
        ''' <param name="top">The Y-pos of the control.</param>
        ''' <param name="width">The width of the control.</param>
        ''' <param name="height">The height of the control.</param>
        ''' <param name="minimum">The progress minimum value</param>
        ''' <param name="maximum">The progress maximum value. Use 0 if the max value is indeterminate.</param>
        ''' <returns>The neame of the ProgressBar</returns>
        <ReturnValueType(VariableType.ProgressBar)>
        <ExMethod>
        Public Shared Function AddProgressBar(
                         formName As Primitive,
                         progressBarName As Primitive,
                         left As Primitive,
                         top As Primitive,
                         width As Primitive,
                         height As Primitive,
                         minimum As Primitive,
                         maximum As Primitive
                    ) As Primitive

            Dim key = ValidateArgs(formName, progressBarName)
            App.Invoke(
                  Sub()
                      Try
                          Dim frm = CType(Forms._forms(CStr(formName).ToLower), System.Windows.Window)

                          Dim pb As New Wpf.ProgressBar With {
                               .Name = progressBarName,
                               .Width = width,
                               .Height = height,
                               .Minimum = minimum,
                               .Maximum = maximum
                          }

                          Wpf.Canvas.SetLeft(pb, left)
                          Wpf.Canvas.SetTop(pb, top)

                          Dim cnv = GetCanvas(frm)
                          cnv.Children.Add(pb)
                          Forms._controls(key) = pb

                      Catch ex As Exception
                          ReportSubError(formName, "AddProgressBar", ex)
                      End Try
                  End Sub)

            Return key
        End Function

        ''' <summary>
        ''' Adds a new Slider control to the form
        ''' </summary>
        ''' <param name="sliderName">A unigue name of the new Slider.</param>
        ''' <param name="left">The X-pos of the control.</param>
        ''' <param name="top">The Y-pos of the control.</param>
        ''' <param name="width">The width of the control.</param>
        ''' <param name="height">The height of the control.</param>
        ''' <param name="minimum">The slider minimum value</param>
        ''' <param name="maximum">The slider maximum value.</param>
        ''' <param name="value">The slider current value</param>
        ''' <param name="tickFrequency">The distance between slide ticks</param>
        ''' <returns>The neame of the Slider</returns>
        <ReturnValueType(VariableType.Slider)>
        <ExMethod>
        Public Shared Function AddSlider(
                         formName As Primitive,
                         sliderName As Primitive,
                         left As Primitive,
                         top As Primitive,
                         width As Primitive,
                         height As Primitive,
                         minimum As Primitive,
                         maximum As Primitive,
                         value As Primitive,
                         tickFrequency As Primitive
                    ) As Primitive

            Dim key = ValidateArgs(formName, sliderName)
            App.Invoke(
                  Sub()
                      Try
                          Dim frm = CType(Forms._forms(CStr(formName).ToLower), System.Windows.Window)

                          Dim s As New Wpf.Slider With {
                               .Name = sliderName,
                               .Width = width,
                               .Height = height,
                               .Minimum = minimum,
                               .Maximum = maximum,
                               .TickFrequency = tickFrequency,
                               .Value = value
                          }

                          Wpf.Canvas.SetLeft(s, left)
                          Wpf.Canvas.SetTop(s, top)

                          Dim cnv = GetCanvas(frm)
                          cnv.Children.Add(s)
                          Forms._controls(key) = s

                      Catch ex As Exception
                          ReportSubError(formName, "AddSlider", ex)
                      End Try
                  End Sub)

            Return key
        End Function

        ''' <summary>
        ''' Adds a new ScrollBar control to the form
        ''' </summary>
        ''' <param name="scrollBarName">A unigue name of the new ScrollBar.</param>
        ''' <param name="left">The X-pos of the control.</param>
        ''' <param name="top">The Y-pos of the control.</param>
        ''' <param name="width">The width of the control.</param>
        ''' <param name="height">The height of the control.</param>
        ''' <param name="minimum">The scrollbar minimum value</param>
        ''' <param name="maximum">The scrollbar maximum value.</param>
        ''' <param name="value">The scrollbar current value</param>
        ''' <returns>The neame of the scrollbar.</returns>
        <ReturnValueType(VariableType.ScrollBar)>
        <ExMethod>
        Public Shared Function AddScrollBar(
                         formName As Primitive,
                         scrollBarName As Primitive,
                         left As Primitive,
                         top As Primitive,
                         width As Primitive,
                         height As Primitive,
                         minimum As Primitive,
                         maximum As Primitive,
                         value As Primitive
                   ) As Primitive

            Dim key = ValidateArgs(formName, scrollBarName)
            App.Invoke(
                  Sub()
                      Try
                          Dim frm = CType(Forms._forms(CStr(formName).ToLower), Window)

                          Dim s As New Wpf.Primitives.ScrollBar With {
                               .Name = scrollBarName,
                               .Width = width,
                               .Height = height,
                               .Minimum = minimum,
                               .Maximum = maximum,
                               .Value = value
                          }

                          Wpf.Canvas.SetLeft(s, left)
                          Wpf.Canvas.SetTop(s, top)

                          Dim cnv = GetCanvas(frm)
                          cnv.Children.Add(s)
                          Forms._controls(key) = s

                      Catch ex As Exception
                          ReportSubError(formName, "AddScrollBar", ex)
                      End Try
                  End Sub)

            Return key
        End Function

        Private Shared ReadOnly ArgsArrProperty As DependencyProperty =
                DependencyProperty.RegisterAttached("ArgsArr",
                GetType(String), GetType(System.Windows.Window)
        )

        ''' <summary>
        ''' Returns the additional data sent to the form via Forms.ShowForm and Forms.ShowDialog methods.
        ''' </summary>
        <ReturnValueType(VariableType.Array)>
        <ExProperty>
        Public Shared Function GetArgsArr(formName As Primitive) As Primitive
            App.Invoke(
                  Sub()
                      Try
                          Dim frm = Forms.GetForm(formName)
                          GetArgsArr = CStr(frm.GetValue(ArgsArrProperty))
                      Catch ex As Exception
                          ReportError(formName, "ArgsArr", ex)
                      End Try
                  End Sub)
        End Function

        Public Shared Sub SetArgsArr(formName As Primitive, value As Primitive)
            App.Invoke(
                   Sub()
                       Try
                           Dim frm = Forms.GetForm(formName)
                           frm.SetValue(ArgsArrProperty, value.AsString())
                       Catch ex As Exception
                           Control.ReportPropertyError(formName, "ArgsArr", value, ex)
                       End Try
                   End Sub)
        End Sub


        ''' <summary>
        ''' Returns True if the form dsiplays a control with the given name.
        ''' </summary>
        ''' <returns>True or False</returns>
        <ReturnValueType(VariableType.Boolean)>
        <ExMethod>
        Public Shared Function ContainsControl(formName As Primitive, controlName As Primitive) As Primitive
            Dim frmName = CStr(formName).ToLower()
            If frmName = "" Then
                Dim msg = "Form name can't be an empty string."
                Dim ex As New ArgumentException(msg)
                Helper.ReportError(msg, ex)
                Throw ex
            End If

            If CStr(controlName) = "" Then
                Dim msg = "Control name can't be an empty string."
                Dim ex As New ArgumentException(msg)
                Helper.ReportError(msg, ex)
                Throw ex
            End If

            Dim key = frmName & "." & controlName.ToString().ToLower()

            Try
                Return Forms._controls.ContainsKey(key)
            Catch ex As Exception
                Helper.ReportError(ex.Message, ex)
                Throw ex
            End Try

        End Function


        ''' <summary>
        ''' Returns an array containg the names of all controls displayed on the form
        ''' </summary>
        <ReturnValueType(VariableType.Array)>
        <ExProperty>
        Public Shared Function GetControls(formName As Primitive) As Primitive
            Dim frmName = formName.ToString().ToLower()
            If Not Forms._forms.ContainsKey(frmName) Then
                ReportSubError(formName, "GetControls", New Exception($"There is no form named `{formName}`"))
                Return ""
            End If

            Dim map = New Dictionary(Of Primitive, Primitive)
            Dim prefix = frmName & "."
            Dim suffex = "." & frmName
            Dim num = 0

            For Each key In Forms._controls.Keys
                If key.StartsWith(prefix) Then
                    map(num) = key
                End If
                num += 1
            Next

            Dim result As New Primitive
            result._arrayMap = map
            Return result
        End Function

        ''' <summary>
        ''' Displayes the form on the screen.
        ''' </summary>
        <ExMethod>
        Public Shared Sub Show(formName As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Dim wnd = Forms.GetForm(formName)
                        wnd.Show()
                        wnd.Activate()
                    Catch ex As Exception
                        ReportSubError(formName, "Show", ex)
                    End Try

                End Sub)
        End Sub

        Private Shared _dialogResult As String

        ''' <summary>
        ''' Displayes the form on the screen as a modal dialog, so the user must close it first to ba able to accees other forms of your app.
        ''' </summary>
        ''' <returns>the dialog result that represnts the type of the button that user clicked, like OK, Yes, No, ... etc.</returns>
        <ReturnValueType(VariableType.DialogResult)>
        <ExMethod>
        Public Shared Function ShowDialog(formName As Primitive) As Primitive
            App.Invoke(
               Sub()
                   Try
                       _dialogResult = ""
                       Dim wnd = Forms.GetForm(formName)
                       Dim dialogResult1 = wnd.ShowDialog()

                       If dialogResult1 Is Nothing Then
                           If _dialogResult = "" Then
                               ShowDialog = DialogResults.No
                           Else
                               ShowDialog = _dialogResult
                           End If

                       ElseIf _dialogResult = "" Then
                           ShowDialog = If(dialogResult1, DialogResults.Yes, DialogResults.Cancel)
                       Else
                           ShowDialog = _dialogResult
                       End If

                   Catch ex As Exception
                       ReportSubError(formName, "ShowDialog", ex)
                   End Try
               End Sub)
        End Function

        ''' <summary>
        ''' Shows the form that has the given name as a child form of the current form.
        ''' </summary>
        ''' <param name="childFormName">the name of the form.</param>
        ''' <param name="argsArr">any additional data, array, or a dynamic object you want to pass to the form. It will be stored in the ArgsArr property of the form, so you can use it as you want</param>
        ''' <returns>the child form name</returns>
        <ReturnValueType(VariableType.Form)>
        <ExMethod>
        Public Shared Function ShowChildForm(parentFormName As Primitive, childFormName As Primitive, argsArr As Primitive) As Primitive
            Dim asm = System.Reflection.Assembly.GetCallingAssembly()
            App.Invoke(
                  Sub()
                      Try
                          Dim parentWnd = Forms.GetForm(parentFormName)

                          If GetIsLoaded(childFormName) Then
                              Dim childWnd = Forms.GetForm(childFormName)
                              childWnd.Owner = parentWnd
                              SetArgsArr(childFormName, argsArr)
                              Show(childFormName)
                              childWnd.RaiseEvent(New RoutedEventArgs(OnFormShownEvent))
                          Else
                              Stack.PushValue("_" & childFormName.AsString().ToLower() & "_argsArr", argsArr)
                              Form.Initialize(childFormName, asm)
                              Dim childWnd = Forms.GetForm(childFormName)
                              childWnd.Owner = parentWnd
                          End If

                      Catch ex As Exception
                          Form.ReportSubError(childFormName, "ShowChildForm", ex)
                      End Try
                  End Sub)

            Return childFormName
        End Function


        ''' <summary>
        ''' Gets or sets the name of the button that the user clicked when he closes the dialog form.
        ''' </summary>
        <ReturnValueType(VariableType.DialogResult)>
        <ExProperty>
        Public Shared Function GetDialogResult(formName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetDialogResult = _dialogResult
                    Catch ex As Exception
                        ReportError(formName, "DialogResult", ex)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetDialogResult(formName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        _dialogResult = value
                    Catch ex As Exception
                        ReportError(formName, "DialogResult", value, ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Shows a message box dialog on the owner form.
        ''' </summary>
        ''' <param name="message">the text to dislpay on the message box</param>
        ''' <param name="title">the title to display of the dialog box</param>
        <ExMethod>
        Public Shared Sub ShowMessage(ownerFormName As Primitive, message As Primitive, title As Primitive)
            App.Invoke(
                Sub()
                    Try
                        System.Windows.MessageBox.Show(Forms.GetForm(ownerFormName), message.ToString(), title.ToString())
                    Catch ex As Exception
                        ReportSubError(ownerFormName, "ShowMessage", ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Hides the form
        ''' </summary>
        <ExMethod>
        Public Shared Sub Hide(formName As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Forms.GetForm(formName).Hide()
                    Catch ex As Exception
                        ReportSubError(formName, "Hide", ex)
                    End Try
                End Sub)
        End Sub


        ''' <summary>
        ''' Closes the form. You can't show the form after it is closed.
        ''' </summary>
        <ExMethod>
        Public Shared Sub Close(formName As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Forms.GetForm(formName).Close()
                    Catch ex As Exception
                        ReportSubError(formName, "Close", ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Returns True if the current form is loaded, regardless if it is hidden in this moment.
        ''' </summary>
        <ReturnValueType(VariableType.Boolean)>
        <ExProperty>
        Public Shared Function GetIsLoaded(formName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetIsLoaded = Forms.GetForm(formName) IsNot Nothing
                    Catch ex As Exception
                        ReportError(formName, "IsLoaded", ex)
                    End Try
                End Sub)
        End Function


        ''' <summary>
        ''' Gets or sets the title of the form
        ''' </summary>
        <ReturnValueType(VariableType.String)>
        <ExProperty>
        Public Shared Function GetText(formName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetText = Forms.GetForm(formName).Title
                    Catch ex As Exception
                        ReportError(formName, "Text", ex)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetText(formName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Forms.GetForm(formName).Title = value
                    Catch ex As Exception
                        ReportError(formName, "Text", value, ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Call this method to allow using transparent colors on the form background.
        ''' </summary>
        <ExMethod>
        Public Shared Sub AllowTransparency(formName As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Dim frm = Forms.GetForm(formName)
                        frm.ResizeMode = ResizeMode.NoResize
                        frm.WindowStyle = WindowStyle.None
                        frm.AllowsTransparency = True
                        frm.Background = Brushes.Transparent
                        Dim canvas = GetCanvas(frm)
                        canvas.Focusable = True
                        AddHandler canvas.PreviewMouseDown, Sub() canvas.Focus()
                    Catch ex As Exception
                        Helper.ReportError(ex.Message, ex)
                    End Try
                End Sub)
        End Sub

        Friend Shared Function GetCanvas(frm As Window) As Wpf.Canvas
            Dim content = frm.Content
            If TypeOf content Is Wpf.Canvas Then Return content
            Dim sp = CType(content, Wpf.StackPanel)
            Return sp.Children(1)
        End Function

        ''' <summary>
        ''' Raises the OnLostFocus event for every control on the current form, to apply any validation logic supplied by you in each OnLostFocuse handler, then checks the Error property to see if the control has errors or not. The process stops at first control with errors and returns false.
        ''' </summary>
        ''' <returns>True if the all controls are valid, or False otherwise.</returns>
        <ExMethod>
        <ReturnValueType(VariableType.Boolean)>
        Public Shared Function Validate(formName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        For Each cntrl In Forms._controls
                            If Control.Validate(cntrl.Key) = False Then
                                Sound.PlayBellRing()
                                cntrl.Value.Focus()
                                Validate = False
                                Return
                            End If
                        Next
                        Validate = True

                    Catch ex As Exception
                        ReportSubError(formName, "Validate", ex)
                    End Try
                End Sub)
        End Function

        ''' <summary>
        ''' Saves the current form design to the given xaml file. This is useful if you use code to create a form and add controls to it in runtime, so you can save it's design to a file if you want.
        ''' </summary>
        ''' <param name="xamlFile">the path an name of the file you want to save the design to.</param>
        <ExMethod>
        Public Sub SaveDesign(formName As Primitive, xamlFile As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Dim xamlPath = IO.Path.GetFullPath(xamlFile.AsString())
                        If IO.Path.GetExtension(xamlPath) = "" Then
                            xamlPath += ".xaml"
                        End If

                        Dim win = Forms.GetForm(formName)
                        Dim canvas = GetCanvas(win)
                        IO.File.WriteAllText(
                            xamlPath,
                            System.Windows.Markup.XamlWriter.Save(canvas)
                        )
                    Catch ex As Exception
                        ReportSubError(formName, "SaveDesign", ex)
                    End Try
                End Sub)
        End Sub

#Region "Events"
        Public Shared ReadOnly OnFormShownEvent As System.Windows.RoutedEvent =
                   System.Windows.EventManager.RegisterRoutedEvent(
                   "OnFormShown",
                   System.Windows.RoutingStrategy.Bubble,
                   GetType(System.Windows.RoutedEventHandler),
                   GetType(System.Windows.Window)
              )

        Public Shared Sub AddAttachedActionHandler(
                     Element As System.Windows.UIElement,
                     Handler As System.Windows.RoutedEventHandler
                )

            If Element IsNot Nothing Then
                Element.AddHandler(OnFormShownEvent, Handler)
            End If
        End Sub

        Public Shared Sub RemoveAttachedActionHandler(
                     Element As System.Windows.UIElement,
                     Handler As System.Windows.RoutedEventHandler
                )

            If Element IsNot Nothing Then
                Element.RemoveHandler(OnFormShownEvent, Handler)
            End If
        End Sub


        ''' <summary>
        ''' Fired after the form is shown and all controls are rendered and are ready to use their properties.
        ''' </summary>
        Public Shared Custom Event OnShown As SmallBasicCallback
            AddHandler(handler As SmallBasicCallback)
                App.Invoke(
                     Sub()
                         Dim form = CType(Forms.GetForm([Event].SenderControl), System.Windows.Window)

                         Try
                             Dim h = Sub(sender As Object, e As EventArgs)
                                         Try
                                             Dim win = CType(sender, System.Windows.Window)
                                             [Event].SenderControl = win.Name & "." & win.Name
                                             Call handler()
                                             [Event].Handled = False


                                         Catch ex As Exception
                                             Helper.ReportError($"The event handler sub `{handler.Method.Name}` fired by the `{NameOf(OnShown)}`event,  caused this error: {ex.Message}", ex)
                                         End Try
                                     End Sub

                             AddHandler form.ContentRendered, h
                             AddAttachedActionHandler(form, h)

                         Catch ex As Exception
                             [Event].ShowErrorMessage(NameOf(OnShown), ex)
                         End Try
                     End Sub)
            End AddHandler

            RemoveHandler(handler As SmallBasicCallback)
            End RemoveHandler

            RaiseEvent()
            End RaiseEvent
        End Event

        ''' <summary>
        ''' Fired jsut before the form is closed.
        ''' Use Event.Handled = True if you want to cancel closing the form.
        ''' </summary>
        Public Shared Custom Event OnClosing As SmallBasicCallback
            AddHandler(handler As SmallBasicCallback)
                App.Invoke(
                     Sub()
                         Dim form = Forms.GetForm([Event].SenderControl)
                         Try
                             AddHandler form.Closing,
                                 Sub(sender As Object, e As ComponentModel.CancelEventArgs)
                                     Try
                                         Dim win = CType(sender, System.Windows.Window)
                                         [Event].SenderControl = win.Name & "." & win.Name

                                         Call handler()

                                         ' the handler may set the Handled property. We will use it and reset it.
                                         e.Cancel = [Event].Handled
                                         [Event].Handled = False
                                     Catch ex As Exception
                                         Helper.ReportError($"The event handler sub `{handler.Method.Name}` fired by the `{NameOf(OnClosing)}`event,  caused this error: {ex.Message}", ex)
                                     End Try
                                 End Sub
                         Catch ex As Exception
                             [Event].ShowErrorMessage(NameOf(OnClosing), ex)
                         End Try
                     End Sub)
            End AddHandler

            RemoveHandler(handler As SmallBasicCallback)
            End RemoveHandler

            RaiseEvent()
            End RaiseEvent
        End Event

        ''' <summary>
        ''' Fired after the form is closed
        ''' </summary>
        Public Shared Custom Event OnClosed As SmallBasicCallback
            AddHandler(handler As SmallBasicCallback)
                App.Invoke(
                     Sub()
                         Dim form = Forms.GetForm([Event].SenderControl)
                         Try
                             AddHandler form.Closed,
                                 Sub(sender As Object, e As EventArgs)
                                     Try
                                         Dim win = CType(sender, System.Windows.Window)
                                         [Event].SenderControl = win.Name & "." & win.Name

                                         Call handler()
                                         [Event].Handled = False
                                     Catch ex As Exception
                                         Helper.ReportError($"The event handler sub `{handler.Method.Name}` fired by the `{NameOf(OnClosing)}`event,  caused this error: {ex.Message}", ex)
                                     End Try
                                 End Sub
                         Catch ex As Exception
                             [Event].ShowErrorMessage(NameOf(OnClosing), ex)
                         End Try
                     End Sub)
            End AddHandler

            RemoveHandler(handler As SmallBasicCallback)
            End RemoveHandler

            RaiseEvent()
            End RaiseEvent
        End Event

#End Region

    End Class
End Namespace