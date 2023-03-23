Imports Microsoft.SmallVisualBasic.Library
Imports Wpf = System.Windows.Controls
Imports App = Microsoft.SmallVisualBasic.Library.Internal.SmallBasicApplication
Imports System.Windows.Media
Imports System.Windows

Namespace WinForms
    ''' <summary>
    ''' Represents the Form control, which is the window that the user interacts with.
    ''' It is easier use the form designer to add new forms and save it to a XAML file.
    ''' It is also possible to crate forms at run time, by calling the Forms.AddForm method to create a new form, or the Forms.LoadForm method to load a form from the XAML file that contains its design, then call the Form.Show method to display the form.
    ''' </summary>
    <SmallVisualBasicType>
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

        ''' <summary>
        ''' Runs the test functions written in the current form, and shows the test results in the TxtTest textbox. If the form doesn't contain a textbox with this name, it will be added at run time to show the results.
        ''' The test function must follow these rules:
        ''' 1. Its name must start with `Test_`, like `Test_FindNames`.
        ''' 2. It must be a function not a sub, and it can't have any parameters.
        ''' 3. The function return value should be a string containing the test result like "passed" or "failed". 
        ''' See the tests written in the "UnitTest Sample" project in the samples folder.
        ''' </summary>
        ''' <returns>the number of tests that have been run.</returns>
        <ExMethod>
        <ReturnValueType(VariableType.Double)>
        Public Shared Function RunTests(formName As Primitive) As Primitive
            Dim asm = System.Reflection.Assembly.GetEntryAssembly()
            RunTests = 0

            App.Invoke(
                Sub()
                    Try
                        Dim frm = formName.AsString().ToLower()
                        Dim frmName = "_SmallVisualBasic_" & frm
                        Dim frmType = asm.GetType(frmName)
                        Dim methods = frmType?.GetMethods(System.Reflection.BindingFlags.Public Or System.Reflection.BindingFlags.Static)
                        If methods Is Nothing Then Return

                        Dim testMethods =
                                From method In methods
                                Where method.Name.ToLower().StartsWith("test_")

                        If Not testMethods.Any Then Return

                        Dim key = frm & ".txttest"
                        If Not Forms._controls.ContainsKey(key) Then
                            Dim wind = CType(Forms._forms(frm), Window)
                            Dim canv = GetCanvas(wind)

                            AddTextBox(
                                    formName,
                                    "txttest",
                                    5,
                                    5,
                                    canv.ActualWidth - 10,
                                    canv.ActualHeight - 10
                            )
                        End If

                        Dim txtTest = CType(Forms._controls(key), Wpf.TextBox)
                        Dim errMsg = " doesn't return a value. Use a test function and return a text showing the result of the test."
                        Dim n = 0
                        txtTest.IsReadOnly = True

                        For Each m In testMethods
                            Dim testName = m.Name
                            Try
                                Dim msg = m.Invoke(Nothing, Nothing)

                                If msg Is Nothing Then
                                    txtTest.AppendText(testName)
                                    txtTest.AppendText(errMsg)
                                Else
                                    txtTest.AppendText(msg.ToString())
                                    n += 1
                                End If

                            Catch ex As Exception
                                txtTest.AppendText($"{testName} has caused the error: {ex.Message}.")
                            End Try
                            txtTest.AppendText(vbCrLf)
                        Next

                        RunTests = n

                    Catch ex As Exception
                        ReportSubError(formName, "RunTests", ex)
                    End Try
                End Sub)
        End Function

        Public Shared Sub Initialize(formName As String, asm As System.Reflection.Assembly)
            App.Invoke(
                Sub()
                    Try
                        Dim frmName = "_SmallVisualBasic_" & formName.ToLower()
                        Dim frm = asm.GetType(frmName)
                        If frm Is Nothing Then
                            frm = System.Reflection.Assembly.GetEntryAssembly.GetType(frmName)
                        End If
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
        ''' <param name="textBoxName">A unique name for the new TextBox.</param>
        ''' <param name="left">The X-pos of the control.</param>
        ''' <param name="top">The Y-pos of the control.</param>
        ''' <param name="width">The width of the control. Use -1 for auto-width</param>
        ''' <param name="height">The height of the control. Use -1 for auto-height</param>
        ''' <returns>
        ''' The key of the TextBox in the formula {formName.textBoxName} like "form1.textbox1".
        ''' sVB can deal with this key as an object of type TextBox, so, you can access the TextBox methods via it.
        ''' </returns>
        <ReturnValueType(VariableType.TextBox)>
        <ExMethod>
        Public Shared Function AddTextBox(
                         formName As Primitive,
                         textBoxName As Primitive,
                         left As Primitive, top As Primitive,
                         width As Primitive, height As Primitive
                    ) As Primitive

            Dim key = ValidateArgs(formName, textBoxName)
            App.Invoke(
                 Sub()
                     Try
                         Dim frm = CType(Forms._forms(CStr(formName).ToLower()), Window)

                         Dim textBox1 As New Wpf.TextBox With {
                           .Name = textBoxName,
                           .Width = GetDouble(width),
                           .Height = GetDouble(height),
                           .VerticalScrollBarVisibility = Wpf.ScrollBarVisibility.Auto,
                           .HorizontalScrollBarVisibility = Wpf.ScrollBarVisibility.Auto
                         }

                         Wpf.Canvas.SetLeft(textBox1, left)
                         Wpf.Canvas.SetTop(textBox1, top)

                         Dim cnv = GetCanvas(frm)
                         cnv.Children.Add(textBox1)
                         Forms._controls(key) = textBox1

                     Catch ex As Exception
                         ReportSubError(formName, "AddTextBox", ex)
                     End Try
                 End Sub)

            Return key
        End Function

        ''' <summary>
        ''' Adds a new Label control to the form
        ''' </summary>
        ''' <param name="labelName">A unique name for the new Label.</param>
        ''' <param name="left">The X-pos of the control.</param>
        ''' <param name="top">The Y-pos of the control.</param>
        ''' <param name="width">The width of the control. Use -1 for auto-width</param>
        ''' <param name="height">The height of the control. Use -1 for auto-height</param>
        ''' <returns>
        ''' The key of the label in the formula {formName.labelName} like "form1.label1".
        ''' sVB can deal with this key as an object of type Label, so, you can access the Label methods via it.
        ''' </returns>
        <ReturnValueType(VariableType.Label)>
        <ExMethod>
        Public Shared Function AddLabel(
                         formName As Primitive,
                         labelName As Primitive,
                         left As Primitive, top As Primitive,
                         width As Primitive, height As Primitive
                   ) As Primitive

            Dim key = ValidateArgs(formName, labelName)
            App.Invoke(
              Sub()
                  Try
                      Dim frm = CType(Forms._forms(CStr(formName).ToLower), System.Windows.Window)

                      Dim label1 As New Wpf.Label With {
                          .Name = labelName,
                          .Width = GetDouble(width),
                          .Height = GetDouble(height)
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
        ''' <param name="imageBoxName">A unique name for the new ImageBox.</param>
        ''' <param name="left">The X-pos of the control.</param>
        ''' <param name="top">The Y-pos of the control.</param>
        ''' <param name="width">The width of the control. Use -1 for auto-width</param>
        ''' <param name="height">The height of the control. Use -1 for auto-height</param>
        ''' <param name="fileName">the path of the image file</param>
        ''' <returns>
        ''' The key of the ImageBox in the formula {formName.imageBoxName} like "form1.imagebox1".
        ''' sVB can deal with this key as an object of type ImageBox, so, you can access the ImageBox methods via it.
        ''' </returns>
        <ReturnValueType(VariableType.ImageBox)>
        <ExMethod>
        Private Shared Function AddImageBox(
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
                                .Width = GetDouble(width),
                                .Height = GetDouble(height)
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
        ''' <param name="timerName">A unique name for the new Timer.</param>
        ''' <param name="interval">The delay time in milliseconds between ticks</param>
        ''' <returns>
        ''' The key of the Timer in the formula {formName.timerName} like "form1.timer1".
        ''' sVB can deal with this key as an object of type WinTimer, so, you can access the WinTimer methods via it.
        ''' </returns>
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


        Friend Shared Function ValidateArgs(formName As String, ByRef controlName As String) As String
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

            Dim frmName = formName.ToLower()
            If Not Forms._forms.ContainsKey(frmName) Then
                Dim msg = $"The Form `{formName}`doesn't exist"
                Dim ex As New ArgumentException(msg)
                Helper.ReportError(msg, ex)
                Throw ex
            End If

            Dim dot = controlName.IndexOf(".")
            If dot > -1 Then
                controlName = controlName.Substring(dot + 1)
            End If
            Dim key = frmName & "." & controlName.ToLower()

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
        ''' <returns>
        ''' The key of the menu in the formula {formName.mainMenuName} like "form1.mainMenu1".
        ''' sVB can deal with this key as an object of type MainMenu, so, you can access the MainMenu methods via it.
        ''' </returns>

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
        ''' <param name="buttonName">A unique name for the new Button.</param>
        ''' <param name="left">The X-pos of the control.</param>
        ''' <param name="top">The Y-pos of the control.</param>
        ''' <param name="width">The width of the control. Use -1 for auto-width</param>
        ''' <param name="height">The height of the control. Use -1 for auto-height</param>
        ''' <returns>
        ''' The key of the button in the formula {formName.buttonName} like "form1.button1".
        ''' sVB can deal with this key as an object of type Button, so, you can access the Button methods via it.
        ''' </returns>
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
                               .Width = GetDouble(width),
                               .Height = GetDouble(height)
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
        ''' <param name="toggleButtonName">A unique name for the new ToggleButton.</param>
        ''' <param name="left">The X-pos of the control.</param>
        ''' <param name="top">The Y-pos of the control.</param>
        ''' <param name="width">The width of the control. Use -1 for auto-width</param>
        ''' <param name="height">The height of the control. Use -1 for auto-height</param>
        ''' <returns>
        ''' The key of the ToggleButton in the formula {formName.toggleButtonName} like "form1.toggleButton1".
        ''' sVB can deal with this key as an object of type ToggleButton, so, you can access the ToggleButton methods via it.
        ''' </returns>
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
                          Dim frm = CType(Forms._forms(CStr(formName).ToLower), Window)

                          Dim toggleButton1 As New Wpf.Primitives.ToggleButton With {
                               .Name = toggleButtonName,
                               .Width = GetDouble(width),
                               .Height = GetDouble(height)
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
        ''' <param name="checkBoxName">A unique name for the new CheckBox.</param>
        ''' <param name="left">The X-pos of the control.</param>
        ''' <param name="top">The Y-pos of the control.</param>
        ''' <param name="text">the text to desply on the CheckBox</param>
        ''' <param name="isChecked">The value to set to the Checked property</param>
        ''' <returns>
        ''' The key of the CheckBox in the formula {formName.checkBoxName} like "form1.checkbox1".
        ''' sVB can deal with this key as an object of type CheckBox, so, you can access the CheckBox methods via it.
        ''' </returns>
        <ReturnValueType(VariableType.CheckBox)>
        <ExMethod>
        Public Shared Function AddCheckBox(
                         formName As Primitive,
                         checkBoxName As Primitive,
                         left As Primitive,
                         top As Primitive,
                         text As Primitive,
                         isChecked As Primitive
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
                          If isChecked.AsString() = "" Then
                              ch.IsChecked = Nothing
                          Else
                              ch.IsChecked = isChecked
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
        ''' <param name="radioButtonName">A unique name for the new RadioButton.</param>
        ''' <param name="left">The X-pos of the control.</param>
        ''' <param name="top">The Y-pos of the control.</param>
        ''' <param name="text">The text to desply on the RadioButton</param>
        ''' <param name="groupName">The name of the group to add the button to</param>
        ''' <param name="isChecked">The value to set to the Checked property</param>
        ''' <returns>
        ''' The key of the RadioButton in the formula {formName.radioButtonName} like "form1.radiobutton1".
        ''' sVB can deal with this key as an object of type RadioButton, so, you can access the RadioButton methods via it.
        ''' </returns>
        <ReturnValueType(VariableType.RadioButton)>
        <ExMethod>
        Public Shared Function AddRadioButton(
                         formName As Primitive,
                         radioButtonName As Primitive,
                         left As Primitive,
                         top As Primitive,
                         text As Primitive,
                         groupName As Primitive,
                         isChecked As Primitive
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
                               .IsChecked = isChecked
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
        ''' <param name="listBoxName">A unique name for the new ListBox.</param>
        ''' <param name="left">The X-pos of the control.</param>
        ''' <param name="top">The Y-pos of the control.</param>
        ''' <param name="width">The width of the control. Use -1 for auto-width</param>
        ''' <param name="height">The height of the control. Use -1 for auto-height</param>
        ''' <returns>
        ''' The key of the listBox in the formula {formName.listBoxName} like "form1.listbox1".
        ''' sVB can deal with this key as an object of type ListBox, so, you can access the ListBox methods via it.
        ''' </returns>
        <ReturnValueType(VariableType.ListBox)>
        <ExMethod>
        Public Shared Function AddListBox(
                         formName As Primitive,
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
                                 .Width = GetDouble(width),
                                 .Height = GetDouble(height)
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
        ''' Adds a new ComboBox control to the form
        ''' </summary>
        ''' <param name="comboBoxName">A unique name for the new ComboBox.</param>
        ''' <param name="left">The X-pos of the control.</param>
        ''' <param name="top">The Y-pos of the control.</param>
        ''' <param name="width">The width of the control. Use -1 for auto-width</param>
        ''' <param name="height">The height of the control. Use -1 for auto-height</param>
        ''' <returns>
        ''' The key of the comboBox in the formula {formName.comboBoxName} like "form1.combobox1".
        ''' sVB can deal with this key as an object of type ComboBox, so, you can use the ComboBox methods via it.
        ''' </returns>
        <ReturnValueType(VariableType.ComboBox)>
        <ExMethod>
        Public Shared Function AddComboBox(
                         formName As Primitive,
                         comboBoxName As Primitive,
                         left As Primitive,
                         top As Primitive,
                         width As Primitive,
                         height As Primitive
                  ) As Primitive

            Dim key = ValidateArgs(formName, comboBoxName)
            App.Invoke(
                  Sub()
                      Try
                          Dim frm = CType(Forms._forms(CStr(formName).ToLower), System.Windows.Window)

                          Dim comboBox1 As New Wpf.ComboBox With {
                                 .Name = comboBoxName,
                                 .Width = GetDouble(width),
                                 .Height = GetDouble(height)
                          }

                          Wpf.Canvas.SetLeft(comboBox1, left)
                          Wpf.Canvas.SetTop(comboBox1, top)

                          Dim cnv = GetCanvas(frm)
                          cnv.Children.Add(comboBox1)
                          Forms._controls(key) = comboBox1

                      Catch ex As Exception
                          ReportSubError(formName, "AddComboBox", ex)
                      End Try
                  End Sub)

            Return key
        End Function

        Private Shared Function GetDouble(value As Primitive) As Double
            Dim v = CDbl(value)
            Return If(v < 0, Double.NaN, v)
        End Function

        ''' <summary>
        ''' Adds a new DatePicker control to the form
        ''' </summary>
        ''' <param name="datePickerName">A unique name for the new DatePicker.</param>
        ''' <param name="left">The X-pos of the control.</param>
        ''' <param name="top">The Y-pos of the control.</param>
        ''' <param name="width">The width of the control. Use -1 for auto-width</param>
        ''' <param name="selectedDate">the date that will be selected in the control</param>
        ''' <returns>
        ''' The key of the DatePicker in the formula {formName.datePikkerName} like "form1.datepicker1".
        ''' sVB can deal with this key as an object of type DatePicker, so, you can access the DatePicker methods via it.
        ''' </returns>
        <ReturnValueType(VariableType.DatePicker)>
        <ExMethod>
        Public Shared Function AddDatePicker(
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
                               .Width = GetDouble(width),
                               .SelectedDate = selectedDate.AsDate()
                          }

                          Wpf.Canvas.SetLeft(dp, left)
                          Wpf.Canvas.SetTop(dp, top)

                          Dim cnv = GetCanvas(frm)
                          cnv.Children.Add(dp)
                          Forms._controls(key) = dp

                      Catch ex As Exception
                          ReportSubError(formName, "AddDatePicker", ex)
                      End Try
                  End Sub)

            Return key
        End Function

        ''' <summary>
        ''' Adds a new ProgressBar control to the form
        ''' </summary>
        ''' <param name="progressBarName">A unique name for the new ProgressBar.</param>
        ''' <param name="left">The X-pos of the control.</param>
        ''' <param name="top">The Y-pos of the control.</param>
        ''' <param name="width">The width of the control. Use -1 for auto-width</param>
        ''' <param name="height">The height of the control. Use -1 for auto-height</param>
        ''' <param name="minimum">The progress minimum value</param>
        ''' <param name="maximum">The progress maximum value. Use 0 if the max value is indeterminate.</param>
        ''' <returns>
        ''' The key of the ProgressBar in the formula {formName.progressBar} like "form1.progressBar1".
        ''' sVB can deal with this key as an object of type ProgressBar, so, you can access the ProgressBar methods via it.
        ''' </returns>
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
                               .Width = GetDouble(width),
                               .Height = GetDouble(height),
                               .Minimum = minimum,
                               .Background = Brushes.White
                          }

                          Wpf.Canvas.SetLeft(pb, left)
                          Wpf.Canvas.SetTop(pb, top)

                          Dim cnv = GetCanvas(frm)
                          cnv.Children.Add(pb)
                          Forms._controls(key) = pb
                          ProgressBar.SetMaximum(key, maximum)

                      Catch ex As Exception
                          ReportSubError(formName, "AddProgressBar", ex)
                      End Try
                  End Sub)

            Return key
        End Function

        ''' <summary>
        ''' Adds a new Slider control to the form
        ''' </summary>
        ''' <param name="sliderName">A unique name for the new Slider.</param>
        ''' <param name="left">The X-pos of the control.</param>
        ''' <param name="top">The Y-pos of the control.</param>
        ''' <param name="width">The width of the control. Use -1 for auto-width</param>
        ''' <param name="height">The height of the control. Use -1 for auto-height</param>
        ''' <param name="minimum">The slider minimum value</param>
        ''' <param name="maximum">The slider maximum value.</param>
        ''' <param name="value">The slider current value</param>
        ''' <param name="tickFrequency">The distance between slide ticks</param>
        ''' <returns>
        ''' The key of the Slider in the formula {formName.sliderName} like "form1.slider1".
        ''' sVB can deal with this key as an object of type Slider, so, you can access the Slider methods via it.
        ''' </returns>
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
                               .Width = GetDouble(width),
                               .Height = GetDouble(height),
                               .Minimum = minimum,
                               .Maximum = maximum,
                               .AutoToolTipPlacement = Wpf.Primitives.AutoToolTipPlacement.TopLeft,
                               .AutoToolTipPrecision = 1,
                               .TickPlacement = Wpf.Primitives.TickPlacement.TopLeft,
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
        ''' <param name="scrollBarName">A unique name for the new ScrollBar.</param>
        ''' <param name="left">The X-pos of the control.</param>
        ''' <param name="top">The Y-pos of the control.</param>
        ''' <param name="width">The width of the control. Use -1 for auto-width</param>
        ''' <param name="height">The height of the control. Use -1 for auto-height</param>
        ''' <param name="minimum">The scrollbar minimum value</param>
        ''' <param name="maximum">The scrollbar maximum value.</param>
        ''' <param name="value">The scrollbar current value</param>
        ''' <returns>
        ''' The key of the Scrollbar in the formula {formName.scrollbarName} like "form1.scrollbar1".
        ''' sVB can deal with this key as an object of type Scrollbar, so, you can access the Scrollbar methods via it.
        ''' </returns>
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
                               .Width = GetDouble(width),
                               .Height = GetDouble(height),
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

        ''' <summary>
        ''' Removes the given control from the current form.
        ''' </summary>
        ''' <param name="controlName">The name or key of the control that you want to remove.</param>
        <ExMethod>
        Public Shared Sub RemoveControl(formName As Primitive, controlName As Primitive)
            Dim key = controlName.AsString().ToLower()
            Dim frm = formName.AsString().ToLower()
            If Not key.StartsWith(frm & ".") Then key = $"{frm}.{controlName}"

            App.Invoke(
                  Sub()
                      If Forms._controls.ContainsKey(key) Then
                          Dim form = CType(Forms._forms(formName), Window)
                          Dim canv = GetCanvas(form)
                          Dim c = Forms._controls(key)

                          If TypeOf c Is Wpf.Menu Then
                              Dim sp = CType(form.Content, Wpf.StackPanel)
                              sp.Children.Clear()
                              form.Content = canv
                          Else
                              canv.Children.Remove(c)
                          End If

                          Forms._controls.Remove(key)
                      End If
                  End Sub)
        End Sub

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
        ''' <param name="controlName">The name of the control to search for. It is case-insensitive.</param>
        ''' <returns>True or False</returns>
        <ReturnValueType(VariableType.Boolean)>
        <ExMethod>
        Public Shared Function ContainsControl(formName As Primitive, controlName As Primitive) As Primitive
            Dim frmName = CStr(formName).ToLower()

            If frmName = "" Then
                Dim msg = "Form name can't be an empty string."
                Dim ex As New ArgumentException(msg)
                Helper.ReportError(msg, ex)
                Return False
            End If

            If CStr(controlName) = "" Then Return False

            Dim key = frmName & "." & controlName.ToString().ToLower()

            Try
                Return Forms._controls.ContainsKey(key)
            Catch ex As Exception
                Helper.ReportError(ex.Message, ex)
            End Try

            Return False
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
            Dim num = 1

            For Each key In Forms._controls.Keys
                If key.StartsWith(prefix) Then
                    map(num) = key
                    num += 1
                End If
            Next

            Dim result As New Primitive
            result._arrayMap = map
            Return result
        End Function

        ''' <summary>
        ''' Gets or sets the image file path to be displayed as an icon on the title bar of the current form.
        ''' </summary>
        <ExProperty>
        Public Shared Function GetIcon(formName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetIcon = CType(Forms.GetForm(formName).Icon, Imaging.BitmapImage).UriSource.AbsolutePath
                    Catch ex As Exception
                        Control.ReportError(formName, "Icon", ex)
                    End Try
                End Sub)
        End Function

        Public Shared Sub SetIcon(formName As Primitive, imageFile As Primitive)
            App.Invoke(
                Sub()
                    Try
                        If Not IO.Path.IsPathRooted(imageFile) Then
                            imageFile = IO.Path.Combine(Program.Directory, imageFile)
                        End If
                        Forms.GetForm(formName).Icon = New Imaging.BitmapImage(New Uri(imageFile))

                    Catch ex As Exception
                        Control.ReportError(formName, "Icon", imageFile, ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Displays the form on the screen, if it is loaded but not shown yet, or if it is hidden.
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
        ''' Displays the form on the screen as a modal dialog, so the user must close it first to ba able to accees other forms of your app.
        ''' </summary>
        ''' <returns>the dialog result that Represents the type of the button that user clicked, like OK, Yes, No, ... etc.</returns>
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
        ''' <param name="childFormName">The name of the child form.</param>
        ''' <param name="argsArr">Any additional data, an array, or a dynamic object you want to pass to the form. It will be stored in the Form.ArgsArr property of the child form, so you can use it as you want.</param>
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
        ''' Closes the form. This will clear any data displayed by its controls, as in fact the form will be totally disposed. You can't call the Form.Show method to show  the form after it is closed, unless you re-created a new instance of it first. If you want to keep the form alive but just hide it, you can just set its Visible property to False to hide it, so you can still be able to show it.
        ''' Note that closing the main form of the project will close the whole application. You can choose the main form by just open any form of the project in sVB and press F5 to run the project, so that form will be the first form displayed when the application starts up.
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
            Dim content = TryCast(frm.Content, UIElement)
            If content Is Nothing Then Return Nothing

            If TypeOf content Is Wpf.Canvas Then Return CType(content, Wpf.Canvas)
            Return content.GetChild(Of Wpf.Canvas)(True)
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
                            Markup.XamlWriter.Save(canvas)
                        )
                    Catch ex As Exception
                        ReportSubError(formName, "SaveDesign", ex)
                    End Try
                End Sub)
        End Sub

#Region "Events"
        Public Shared ReadOnly OnFormShownEvent As RoutedEvent =
            EventManager.RegisterRoutedEvent(
            "OnFormShown",
            RoutingStrategy.Bubble,
            GetType(RoutedEventHandler),
            GetType(Window)
        )

        Public Shared Sub AddAttachedActionHandler(
                     Element As UIElement,
                     Handler As RoutedEventHandler
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
        Public Shared Custom Event OnShown As SmallVisualBasicCallback
            AddHandler(handler As SmallVisualBasicCallback)
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

            RemoveHandler(handler As SmallVisualBasicCallback)
            End RemoveHandler

            RaiseEvent()
            End RaiseEvent
        End Event

        ''' <summary>
        ''' Fired jsut before the form is closed.
        ''' Use Event.Handled = True if you want to cancel closing the form.
        ''' </summary>
        Public Shared Custom Event OnClosing As SmallVisualBasicCallback
            AddHandler(handler As SmallVisualBasicCallback)
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

            RemoveHandler(handler As SmallVisualBasicCallback)
            End RemoveHandler

            RaiseEvent()
            End RaiseEvent
        End Event

        ''' <summary>
        ''' Fired after the form is closed
        ''' </summary>
        Public Shared Custom Event OnClosed As SmallVisualBasicCallback
            AddHandler(handler As SmallVisualBasicCallback)
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

            RemoveHandler(handler As SmallVisualBasicCallback)
            End RemoveHandler

            RaiseEvent()
            End RaiseEvent
        End Event

#End Region

    End Class
End Namespace