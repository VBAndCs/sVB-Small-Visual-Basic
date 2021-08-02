Imports Microsoft.SmallBasic.Library
Imports Wpf = System.Windows.Controls
Imports App = Microsoft.SmallBasic.Library.Internal.SmallBasicApplication

Namespace WinForms
    <SmallBasicType>
    Public NotInheritable Class Form

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

            app.Invoke(
            Sub()
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

            End Sub)
        End Sub

        <ExMethod>
        Public Shared Function ContainsControl(formName As Primitive, controlName As Primitive) As Primitive
            Dim frmName = CStr(formName).ToLower()
            If frmName = "" Then
                MsgBox("Form name can't be an empty string.")
            End If

            Dim cntrName = CStr(controlName).ToLower()
            If cntrName = "" Then
                MsgBox("Control name can't be an empty string.")
            End If

            Return Forms._forms.ContainsKey(frmName) AndAlso
                    Forms._forms(frmName).ContainsKey(cntrName)
        End Function

        <ExMethod>
        Public Shared Function GetControls(formName As Primitive) As Primitive
            If Not Forms._forms.ContainsKey(formName) Then
                MsgBox($"There is no form names `{formName}`")
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
            'TextWindow.WriteLine($"Showing {formName}")


            App.Invoke(
                Sub()
                    Dim wnd = Forms.GetForm(formName)
                    wnd.Show()
                    wnd.Activate()
                End Sub)
        End Sub

        <ExMethod>
        Public Shared Function ShowDialog(formName As Primitive) As Primitive
            App.Invoke(
               Sub()
                   Dim wnd = Forms.GetForm(formName)
                   ShowDialog = wnd.ShowDialog().GetValueOrDefault
               End Sub)
        End Function

        <ExMethod>
        Public Shared Sub ShowMessage(ownerFormName As Primitive, message As Primitive, title As Primitive)
            App.Invoke(
            Sub() System.Windows.MessageBox.Show(Forms.GetForm(ownerFormName), message.ToString(), title.ToString()))
        End Sub

        <ExMethod>
        Public Shared Sub Hide(formName As Primitive)
            App.Invoke(Sub() Forms.GetForm(formName).Hide())
        End Sub

        <ExMethod>
        Public Shared Sub Close(formName As Primitive)
            App.Invoke(Sub() Forms.GetForm(formName).Close())
        End Sub

        <ExProperty>
        Public Shared Function GetText(formName As Primitive) As Primitive
            App.Invoke(Sub() GetText = Forms.GetForm(formName).Title.ToString())
        End Function

        <ExProperty>
        Public Shared Sub SetText(formName As Primitive, value As Primitive)
            App.Invoke(Sub() Forms.GetForm(formName).Title = value)
        End Sub
    End Class
End Namespace