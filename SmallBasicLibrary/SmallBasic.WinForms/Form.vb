Imports Microsoft.SmallBasic.Library
Imports Wpf = System.Windows.Controls
Imports App = Microsoft.SmallBasic.Library.Internal.SmallBasicApplication
Imports System.Windows.Media

Namespace WinForms
    <SmallBasicType>
    Public NotInheritable Class Form

        Shared Sub ShowSubError(formName As String, memberName As String, msg As String)
            MsgBox($"Calling {formName}.{memberName} caused an error: {vbCrLf}{msg}")
        End Sub


        Shared Sub ShowErrorMesssage(formName As String, memberName As String, msg As String)
            MsgBox($"Reading {formName}.{memberName} caused an error: {vbCrLf}{msg}")
        End Sub

        Shared Sub ShowErrorMesssage(formName As String, memberName As String, value As String, msg As String)
            MsgBox($"Sending `{value}` to {formName}.{memberName} caused an error: {vbCrLf}{msg}")
        End Sub


        ''' <summary>
        ''' Adds a new TextBox control to the form
        ''' </summary>
        ''' <param name="formName">The name of the form. Omit this param if you call this method from the form name</param>
        ''' <param name="textBoxName">A unique name for the control.</param>
        ''' <param name="left">The X-pos of the control.</param>
        ''' <param name="top">The Y-pos of the control.</param>
        ''' <param name="width">The width of the control.</param>
        ''' <param name="height">The height of the control.</param>
        <ExMethod>
        Public Shared Sub AddTexBox(formName As Primitive,
                         textBoxName As Primitive,
                         left As Primitive, top As Primitive,
                         width As Primitive, height As Primitive)

            App.Invoke(
                 Sub()
                     Try
                         Dim frm = Forms.GetForm(formName)

                         If ContainsControl(formName, textBoxName) Then
                             Return
                             MsgBox($"There is another control with the name '{textBoxName}' on the ''{formName} form")
                         End If

                         Dim textBox1 As New Wpf.TextBox With {
                           .Name = textBoxName,
                           .Width = width,
                           .Height = height
                         }

                         Wpf.Canvas.SetLeft(textBox1, left)
                         Wpf.Canvas.SetTop(textBox1, top)

                         Dim cnv As Wpf.Canvas = frm.Content
                         cnv.Children.Add(textBox1)
                         AddToDictionary(formName, textBox1)
                     Catch ex As Exception
                         ShowSubError(formName, "AddTexBox", ex.Message)
                     End Try
                 End Sub)
        End Sub

        ''' <summary>
        ''' Adds a new Label control to the form
        ''' </summary>
        ''' <param name="formName">The name of the form. Omit this param if you call this method from the form name</param>
        ''' <param name="labelName">A unique name for the control.</param>
        ''' <param name="left">The X-pos of the control.</param>
        ''' <param name="top">The Y-pos of the control.</param>
        ''' <param name="width">The width of the control.</param>
        ''' <param name="height">The height of the control.</param>
        <ExMethod>
        Public Shared Sub AddLabel(formName As Primitive,
                         labelName As Primitive,
                         left As Primitive, top As Primitive,
                         width As Primitive, height As Primitive)

            App.Invoke(
                    Sub()
                        Try
                            Dim frm = Forms.GetForm(formName)

                            If ContainsControl(formName, labelName) Then
                                Return
                                MsgBox($"There is another control with the name '{labelName}' on the ''{formName} form")
                            End If

                            Dim label1 As New Wpf.Label With {
                               .Name = labelName,
                               .Width = width,
                               .Height = height
                            }

                            Wpf.Canvas.SetLeft(label1, left)
                            Wpf.Canvas.SetTop(label1, top)

                            Dim cnv As Wpf.Canvas = frm.Content
                            cnv.Children.Add(label1)
                            AddToDictionary(formName, label1)
                        Catch ex As Exception
                            ShowSubError(formName, "AddLabel", ex.Message)
                        End Try
                    End Sub)
        End Sub

        ''' <summary>
        ''' Adds a new Button control to the form
        ''' </summary>
        ''' <param name="formName">The name of the form. Omit this param if you call this method from the form name</param>
        ''' <param name="buttonName">A unique name for the control.</param>
        ''' <param name="left">The X-pos of the control.</param>
        ''' <param name="top">The Y-pos of the control.</param>
        ''' <param name="width">The width of the control.</param>
        ''' <param name="height">The height of the control.</param>
        <ExMethod>
        Public Shared Sub AddButton(formName As Primitive,
                         buttonName As Primitive,
                         left As Primitive, top As Primitive,
                         width As Primitive, height As Primitive)

            App.Invoke(
                  Sub()
                      Try
                          Dim frm = Forms.GetForm(formName)

                          If ContainsControl(formName, buttonName) Then
                              Return
                              MsgBox($"There is another control with the name '{buttonName}' on the ''{formName} form")
                          End If

                          Dim button1 As New Wpf.Button With {
                               .Name = buttonName,
                               .Width = width,
                               .Height = height
                          }

                          Wpf.Canvas.SetLeft(button1, left)
                          Wpf.Canvas.SetTop(button1, top)

                          Dim cnv As Wpf.Canvas = frm.Content
                          cnv.Children.Add(button1)

                          AddToDictionary(formName, button1)
                      Catch ex As Exception
                          ShowSubError(formName, "AddButton", ex.Message)
                      End Try
                  End Sub)
        End Sub

        Private Shared Sub AddToDictionary(FormKey As String, control As Wpf.Control)
            Forms._forms(FormKey.ToLower()).Add(control.Name.ToLower(), control)
        End Sub

        ''' <summary>
        ''' Adds a new Button control to the form
        ''' </summary>
        ''' <param name="formName">The name of the form. Omit this param if you call this method from the form name</param>
        ''' <param name="listBoxName">A unique name for the control.</param>
        ''' <param name="left">The X-pos of the control.</param>
        ''' <param name="top">The Y-pos of the control.</param>
        ''' <param name="width">The width of the control.</param>
        ''' <param name="height">The height of the control.</param>
        <ExMethod>
        Public Shared Sub AddListBox(formName As Primitive,
                         listBoxName As Primitive,
                         left As Primitive, top As Primitive,
                         width As Primitive, height As Primitive)

            App.Invoke(
                  Sub()
                      Try
                          Dim frm = Forms.GetForm(formName)

                          If ContainsControl(formName, listBoxName) Then
                              Return
                              MsgBox($"There is another control with the name '{listBoxName}' on the ''{formName} form")
                          End If

                          Dim listBox1 As New Wpf.ListBox With {
                                 .Name = listBoxName,
                                 .Width = width,
                                 .Height = height
                          }

                          Wpf.Canvas.SetLeft(listBox1, left)
                          Wpf.Canvas.SetTop(listBox1, top)

                          Dim cnv As Wpf.Canvas = frm.Content
                          cnv.Children.Add(listBox1)
                          AddToDictionary(formName, listBox1)
                      Catch ex As Exception
                          ShowSubError(formName, "AddListBox", ex.Message)
                      End Try
                  End Sub)
        End Sub

        <ExMethod>
        Public Shared Function ContainsControl(formName As Primitive, controlName As Primitive) As Primitive
            Dim frmName = CStr(formName).ToLower()
            If frmName = "" Then
                Dim msg = "Form name can't be an empty string."
                MsgBox(msg)
                Throw New Exception(msg)
            End If

            Dim cntrName = CStr(controlName).ToLower()
            If cntrName = "" Then
                Dim msg = "Control name can't be an empty string."
                MsgBox(msg)
                Throw New Exception(msg)
            End If

            Try
                Return Forms._forms.ContainsKey(frmName) AndAlso
                        Forms._forms(frmName).ContainsKey(cntrName)
            Catch ex As Exception
                MsgBox(ex.Message)
                Throw ex
            End Try

        End Function

        <ExMethod>
        Public Shared Function GetControls(formName As Primitive) As Primitive
            If Not Forms._forms.ContainsKey(formName) Then
                ShowSubError(formName, "GetControls", $"There is no form named `{formName}`")
                Return ""
            End If

            Dim map = New Dictionary(Of Primitive, Primitive)
            Dim num = 0
            For Each key In Forms._forms(formName).Keys
                If num > 0 Then map(num) = key ' The first key is the form itself
                num += 1
            Next
            Return Primitive.ConvertFromMap(map)
        End Function


        <ExMethod>
        Public Shared Sub Show(formName As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Dim wnd = Forms.GetForm(formName)
                        wnd.Show()
                        wnd.Activate()
                    Catch ex As Exception
                        ShowSubError(formName, "Show", ex.Message)
                    End Try

                End Sub)
        End Sub

        <ExMethod>
        Public Shared Function ShowDialog(formName As Primitive) As Primitive
            App.Invoke(
               Sub()
                   Try
                       Dim wnd = Forms.GetForm(formName)
                       ShowDialog = wnd.ShowDialog().GetValueOrDefault
                   Catch ex As Exception
                       ShowSubError(formName, "ShowDialog", ex.Message)
                   End Try
               End Sub)
        End Function

        <ExMethod>
        Public Shared Sub ShowMessage(ownerFormName As Primitive, message As Primitive, title As Primitive)
            App.Invoke(
                Sub()
                    Try
                        System.Windows.MessageBox.Show(Forms.GetForm(ownerFormName), message.ToString(), title.ToString())
                    Catch ex As Exception
                        ShowSubError(ownerFormName, "ShowMessage", ex.Message)
                    End Try
                End Sub)
        End Sub

        <ExMethod>
        Public Shared Sub Hide(formName As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Forms.GetForm(formName).Hide()
                    Catch ex As Exception
                        ShowSubError(formName, "Hide", ex.Message)
                    End Try
                End Sub)
        End Sub

        <ExMethod>
        Public Shared Sub Close(formName As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Forms.GetForm(formName).Close()
                    Catch ex As Exception
                        ShowSubError(formName, "Close", ex.Message)
                    End Try
                End Sub)
        End Sub

        <ExProperty>
        Public Shared Function GetText(formName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetText = Forms.GetForm(formName).Title
                    Catch ex As Exception
                        ShowErrorMesssage(formName, "Text", ex.Message)
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
                        ShowErrorMesssage(formName, "Text", value, ex.Message)
                    End Try
                End Sub)
        End Sub

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
                        MsgBox(ex.Message)
                    End Try
                End Sub)
        End Sub


        Public Shared Custom Event OnClosing As SmallBasicCallback
            AddHandler(handler As SmallBasicCallback)
                Dim VisualElement = CType(Control.GetControl([Event].SenderForm, ""), System.Windows.Window)
                Try
                    AddHandler VisualElement.Closing,
                        Sub(sender As Object, e As ComponentModel.CancelEventArgs)
                            Try
                                Dim win = CType(sender, System.Windows.Window)
                                [Event].SenderControl = win.Name
                                [Event].SenderForm = win.Name

                                Call handler()

                                ' the handler may set the Handled property. We will use it and reset it.
                                e.Cancel = [Event].Handled
                                [Event].Handled = False
                            Catch ex As Exception
                                MsgBox($"The event handler sub `{handler.Method.Name}` fired by the `{[Event].SenderForm}.{NameOf(OnClosing)}`event,  caused this error: {ex.Message}")
                            End Try
                        End Sub
                Catch ex As Exception
                    [Event].ShowErrorMessage(NameOf(OnClosing), ex.Message)
                End Try

            End AddHandler

            RemoveHandler(handler As SmallBasicCallback)
            End RemoveHandler

            RaiseEvent()
            End RaiseEvent
        End Event

        Public Shared Custom Event OnClosed As SmallBasicCallback
            AddHandler(handler As SmallBasicCallback)
                Dim VisualElement = CType(Control.GetControl([Event].SenderForm, ""), System.Windows.Window)
                Try
                    AddHandler VisualElement.Closed,
                        Sub(sender As Object, e As EventArgs)
                            Try
                                Dim win = CType(sender, System.Windows.Window)
                                [Event].SenderControl = win.Name
                                [Event].SenderForm = win.Name

                                Call handler()
                                [Event].Handled = False
                            Catch ex As Exception
                                MsgBox($"The event handler sub `{handler.Method.Name}` fired by the `{[Event].SenderForm}.{NameOf(OnClosing)}`event,  caused this error: {ex.Message}")
                            End Try
                        End Sub
                Catch ex As Exception
                    [Event].ShowErrorMessage(NameOf(OnClosing), ex.Message)
                End Try

            End AddHandler

            RemoveHandler(handler As SmallBasicCallback)
            End RemoveHandler

            RaiseEvent()
            End RaiseEvent
        End Event
    End Class
End Namespace