Imports Microsoft.SmallBasic.Library
Imports Wpf = System.Windows.Controls

<SmallBasicType>
Public NotInheritable Class Form
    Public Shared Sub AddTexBox(formName As Primitive,
                         TextBoxName As Primitive,
                         left As Primitive, top As Primitive,
                         width As Primitive, height As Primitive)

        Forms.Dispatcher.Invoke(
            Sub()
                Dim frm = Forms.GetForm(formName)

                If ContainsControl(formName, TextBoxName) Then
                    Throw New ArgumentException($"There is another control with the name '{TextBoxName}' on the ''{formName} form")
                End If

                Dim textBox1 As New Wpf.TextBox With {
                       .Name = TextBoxName,
                       .Width = width,
                       .Height = height
                }

                Canvas.SetLeft(textBox1, left)
                Canvas.SetTop(textBox1, top)

                Dim cnv As Canvas = frm.Content
                cnv.Children.Add(textBox1)
                Forms._forms(formName).Add(TextBoxName, textBox1)

            End Sub)
    End Sub

    Public Shared Sub AddLabel(formName As Primitive,
                         labelName As Primitive,
                         left As Primitive, top As Primitive,
                         width As Primitive, height As Primitive)

        Forms.Dispatcher.Invoke(
            Sub()
                Dim frm = Forms.GetForm(formName)

                If ContainsControl(formName, labelName) Then
                    Throw New ArgumentException($"There is another control with the name '{labelName}' on the ''{formName} form")
                End If

                Dim label1 As New Wpf.Label With {
                       .Name = labelName,
                       .Width = width,
                       .Height = height
                }

                Canvas.SetLeft(label1, left)
                Canvas.SetTop(label1, top)

                Dim cnv As Canvas = frm.Content
                cnv.Children.Add(label1)
                Forms._forms(formName).Add(labelName, label1)

            End Sub)
    End Sub

    Public Shared Sub AddButton(formName As Primitive,
                         buttonName As Primitive,
                         left As Primitive, top As Primitive,
                         width As Primitive, height As Primitive)

        Forms.Dispatcher.Invoke(
            Sub()
                Dim frm = Forms.GetForm(formName)

                If ContainsControl(formName, buttonName) Then
                    Throw New ArgumentException($"There is another control with the name '{buttonName}' on the ''{formName} form")
                End If

                Dim button1 As New Wpf.Button With {
                     .Name = buttonName,
                     .Width = width,
                     .Height = height
                }

                Canvas.SetLeft(button1, left)
                Canvas.SetTop(button1, top)

                Dim cnv As Canvas = frm.Content
                cnv.Children.Add(button1)
                Forms._forms(formName).Add(buttonName, button1)

            End Sub)
    End Sub

    Public Shared Sub AddListBox(formName As Primitive,
                         listBoxName As Primitive,
                         left As Primitive, top As Primitive,
                         width As Primitive, height As Primitive)

        Forms.Dispatcher.Invoke(
            Sub()
                Dim frm = Forms.GetForm(formName)

                If ContainsControl(formName, listBoxName) Then
                    Throw New ArgumentException($"There is another control with the name '{listBoxName}' on the ''{formName} form")
                End If

                Dim listBox1 As New Wpf.ListBox With {
                       .Name = listBoxName,
                       .Width = width,
                       .Height = height
                }

                Canvas.SetLeft(listBox1, left)
                Canvas.SetTop(listBox1, top)

                Dim cnv As Canvas = frm.Content
                cnv.Children.Add(listBox1)
                Forms._forms(formName).Add(listBoxName, listBox1)

            End Sub)
    End Sub

    Public Shared Function ContainsControl(formName As Primitive, controlName As Primitive) As Primitive
        Dim frmName = CStr(formName)
        If frmName = "" Then
            Throw New ArgumentException("Form name can't be an empty string.")
        End If

        Dim cntrName = CStr(controlName)
        If cntrName = "" Then
            Throw New ArgumentException("Control name can't be an empty string.")
        End If

        Return Forms._forms.ContainsKey(frmName) AndAlso
                    Forms._forms(formName).ContainsKey(cntrName)
    End Function

    Public Shared Function GetControls(formName As Primitive) As Primitive
        If Not Forms._forms.ContainsKey(formName) Then
            Throw New ArgumentException($"There is no form names `{formName}`")
        End If

        Dim map = New Dictionary(Of Primitive, Primitive)
        Dim num = 0
        For Each key In Forms._forms(formName).Keys
            If num > 0 Then map(num) = key ' The first key is the form itself
            num += 1
        Next
        Return Primitive.ConvertFromMap(map)
    End Function

    Public Shared Sub Show(formName As Primitive)
        TextWindow.WriteLine("Showing " & formName.ToString())
        Forms.Dispatcher.Invoke(Sub() Forms.GetForm(formName).Show())
    End Sub

    Public Shared Function ShowDialog(formName As Primitive) As Primitive
        Forms.Dispatcher.Invoke(
               Sub()
                   Dim wnd = Forms.GetForm(formName)
                   ShowDialog = wnd.ShowDialog().GetValueOrDefault
               End Sub)
    End Function

    Public Shared Sub Hide(formName As Primitive)
        Forms.Dispatcher.Invoke(Sub() Forms.GetForm(formName).Hide())
    End Sub

    Public Shared Sub Close(formName As Primitive)
        Forms.Dispatcher.Invoke(Sub() Forms.GetForm(formName).Close())
    End Sub


    Public Shared Function GetText(formName As Primitive, __ As Primitive) As Primitive
        Forms.Dispatcher.Invoke(Sub() GetText = Forms.GetForm(formName).Title.ToString())
    End Function

    Public Shared Sub SetText(formName As Primitive, value As Primitive)
        Forms.Dispatcher.Invoke(Sub() Forms.GetForm(formName).Title = value)
    End Sub
End Class
