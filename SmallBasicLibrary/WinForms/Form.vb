Imports Microsoft.SmallVisualBasic.Library
Imports Wpf = System.Windows.Controls
Imports App = Microsoft.SmallVisualBasic.Library.Internal.SmallBasicApplication
Imports System.Windows.Media

Namespace WinForms

    <SmallBasicType>
    <HideFromIntellisense>
    Public NotInheritable Class Form

        Shared Sub ShowSubError(formName As String, memberName As String, ex As Exception)
            ReportError($"Calling {formName}.{memberName} caused an error: {vbCrLf}{ex.Message}", ex)
        End Sub

        Shared Sub ShowErrorMesssage(formName As String, memberName As String, ex As Exception)
            ReportError($"Reading {formName}.{memberName} caused an error: {vbCrLf}{ex.Message}", ex)
        End Sub

        Shared Sub ShowErrorMesssage(formName As String, memberName As String, value As String, ex As Exception)
            ReportError($"Sending `{value}` to {formName}.{memberName} caused an error: {vbCrLf}{ex.Message}", ex)
        End Sub

        Public Shared Sub Initialize(formName As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Dim frmName = "_SmallVisualBasic_" & formName.AsString().ToLower()
                        Dim asm = System.Reflection.Assembly.GetEntryAssembly()
                        Dim frm = asm.GetType(frmName)
                        Dim initializeMethod = frm?.GetMethod("Initialize", System.Reflection.BindingFlags.Public Or System.Reflection.BindingFlags.Static)
                        initializeMethod?.Invoke(Nothing, Nothing)

                    Catch ex As Exception
                        ShowSubError(formName, "Initialize", ex)
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
                         Dim frm = CType(Forms._forms(CStr(formName).ToLower), System.Windows.Window)

                         Dim textBox1 As New Wpf.TextBox With {
                           .Name = textBoxName,
                           .Width = width,
                           .Height = height
                         }

                         Wpf.Canvas.SetLeft(textBox1, left)
                         Wpf.Canvas.SetTop(textBox1, top)

                         Dim cnv As Wpf.Canvas = frm.Content
                         cnv.Children.Add(textBox1)
                         Forms._controls(key) = textBox1

                     Catch ex As Exception
                         ShowSubError(formName, "AddTexBox", ex)
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

                            Dim cnv As Wpf.Canvas = frm.Content
                            cnv.Children.Add(label1)
                            Forms._controls(key) = label1

                        Catch ex As Exception
                            ShowSubError(formName, "AddLabel", ex)
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

                            Dim cnv As Wpf.Canvas = frm.Content
                            cnv.Children.Add(img)
                            Forms._controls(key) = img

                        Catch ex As Exception
                            ShowSubError(formName, "AddImageBox", ex)
                        End Try
                    End Sub)

            Return key
        End Function

        Private Shared Function ValidateArgs(formName As String, controlName As String) As String
            If formName = "" Then
                Dim msg = "`formName` can't be empty string"
                Dim ex As New ArgumentException(msg)
                ReportError(msg, ex)
                Throw ex
            End If

            If controlName = "" Then
                Dim msg = "`buttonName` can't be empty string"
                Dim ex As New ArgumentException(msg)
                ReportError(msg, ex)
                Throw ex
            End If

            Dim frmName = formName.ToString().ToLower
            If Not Forms._forms.ContainsKey(frmName) Then
                Dim msg = $"The Form `{formName}`doesn't exist"
                Dim ex As New ArgumentException(msg)
                ReportError(msg, ex)
                Throw ex
            End If

            Dim key = frmName & "." & controlName.ToString().ToLower()
            If Forms._controls.ContainsKey(key) Then
                Dim msg = $"There is another control with the name '{controlName}' on the ''{formName} form"
                Dim ex As New ArgumentException(msg)
                ReportError(msg, ex)
                Throw ex
            End If

            Return key
        End Function

        ''' <summary>
        ''' Adds a new Button control to the form
        ''' </summary>
        ''' <param name="datePickerName">A unigue name of the new Button.</param>
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

                          Dim cnv As Wpf.Canvas = frm.Content
                          cnv.Children.Add(button1)
                          Forms._controls(key) = button1

                      Catch ex As Exception
                          ShowSubError(formName, "AddButton", ex)
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

                          Dim cnv As Wpf.Canvas = frm.Content
                          cnv.Children.Add(listBox1)
                          Forms._controls(key) = listBox1

                      Catch ex As Exception
                          ShowSubError(formName, "AddListBox", ex)
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
        ''' <returns>The neame of the DatePicker</returns>
        <ReturnValueType(VariableType.DatePicker)>
        <ExMethod>
        Public Shared Function AddDatePiker(
                         formName As Primitive,
                         datePickerName As Primitive,
                         left As Primitive,
                         top As Primitive,
                         width As Primitive
                    ) As Primitive

            Dim key = ValidateArgs(formName, datePickerName)
            App.Invoke(
                  Sub()
                      Try
                          Dim frm = CType(Forms._forms(CStr(formName).ToLower), System.Windows.Window)

                          Dim dp As New Wpf.DatePicker With {
                               .Name = datePickerName,
                               .Width = width
                          }

                          Wpf.Canvas.SetLeft(dp, left)
                          Wpf.Canvas.SetTop(dp, top)

                          Dim cnv As Wpf.Canvas = frm.Content
                          cnv.Children.Add(dp)
                          Forms._controls(key) = dp

                      Catch ex As Exception
                          ShowSubError(formName, "AddDatePiker", ex)
                      End Try
                  End Sub)

            Return key
        End Function

        Private Shared ReadOnly ArgsArrProperty As _
                System.Windows.DependencyProperty =
                       System.Windows.DependencyProperty.RegisterAttached("ArgsArr",
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
                          ShowErrorMesssage(formName, "ArgsArr", ex)
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
                           Control.ShowPropertyMesssage(formName, "ArgsArr", value, ex)
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
                ReportError(msg, ex)
                Throw ex
            End If

            If CStr(controlName) = "" Then
                Dim msg = "Control name can't be an empty string."
                Dim ex As New ArgumentException(msg)
                ReportError(msg, ex)
                Throw ex
            End If

            Dim key = frmName & "." & controlName.ToString().ToLower()

            Try
                Return Forms._controls.ContainsKey(key)
            Catch ex As Exception
                ReportError(ex.Message, ex)
                Throw ex
            End Try

        End Function


        ''' <summary>
        ''' Returns an array containg the names of all controls displayed on the form
        ''' </summary>
        <ReturnValueType(VariableType.Array)>
        <ExMethod>
        Public Shared Function GetControls(formName As Primitive) As Primitive
            If Not Forms._forms.ContainsKey(formName) Then
                ShowSubError(formName, "GetControls", New Exception($"There is no form named `{formName}`"))
                Return ""
            End If

            Dim map = New Dictionary(Of Primitive, Primitive)
            Dim frmName = formName.ToString().ToLower()
            Dim prefix = frmName & "."
            Dim suffex = "." & frmName
            Dim controls = Forms._controls
            Dim num = 0

            For Each key In Forms._controls.Keys
                If key.StartsWith(prefix) AndAlso Not key.EndsWith(suffex) Then
                    map(num) = key
                End If
                num += 1
            Next
            Return Primitive.ConvertFromMap(map)
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
                        ShowSubError(formName, "Show", ex)
                    End Try

                End Sub)
        End Sub

        Private Shared dialogResult2 As String

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
                       dialogResult2 = ""
                       Dim wnd = Forms.GetForm(formName)
                       Dim dialogResult1 = wnd.ShowDialog()

                       If dialogResult1 Is Nothing Then
                           If dialogResult2 = "" Then
                               ShowDialog = DialogResults.No
                           Else
                               ShowDialog = dialogResult2
                           End If
                       ElseIf dialogResult2 = "" Then
                           ShowDialog = If(dialogResult1, DialogResults.Yes, DialogResults.Cancel)
                       Else
                           ShowDialog = dialogResult2
                       End If

                   Catch ex As Exception
                       ShowSubError(formName, "ShowDialog", ex)
                   End Try
               End Sub)
        End Function

        ''' <summary>
        ''' Gets or sets the name of the button that the user clicked when he closes the dialog form.
        ''' </summary>
        <ReturnValueType(VariableType.String)>
        <ExProperty>
        Public Shared Function GetDialogResult(formName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetDialogResult = dialogResult2
                    Catch ex As Exception
                        ShowErrorMesssage(formName, "DialogResult", ex)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetDialogResult(formName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        dialogResult2 = value
                    Catch ex As Exception
                        ShowErrorMesssage(formName, "DialogResult", value, ex)
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
                        ShowSubError(ownerFormName, "ShowMessage", ex)
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
                        ShowSubError(formName, "Hide", ex)
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
                        ShowSubError(formName, "Close", ex)
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
                        ShowErrorMesssage(formName, "IsLoaded", ex)
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
                        ShowErrorMesssage(formName, "Text", ex)
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
                        ShowErrorMesssage(formName, "Text", value, ex)
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
                        frm.ResizeMode = System.Windows.ResizeMode.NoResize
                        frm.WindowStyle = System.Windows.WindowStyle.None
                        frm.AllowsTransparency = True
                        frm.Background = Brushes.Transparent
                        Dim canvas = CType(frm.Content, Wpf.Canvas)
                        canvas.Focusable = True
                        AddHandler canvas.PreviewMouseDown, Sub() canvas.Focus()
                    Catch ex As Exception
                        ReportError(ex.Message, ex)
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
                                             ReportError($"The event handler sub `{handler.Method.Name}` fired by the `{NameOf(OnShown)}`event,  caused this error: {ex.Message}", ex)
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
                                         ReportError($"The event handler sub `{handler.Method.Name}` fired by the `{NameOf(OnClosing)}`event,  caused this error: {ex.Message}", ex)
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
                                         ReportError($"The event handler sub `{handler.Method.Name}` fired by the `{NameOf(OnClosing)}`event,  caused this error: {ex.Message}", ex)
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