Imports Microsoft.SmallBasic.Library
Imports Wpf = System.Windows.Controls
Imports App = Microsoft.SmallBasic.Library.Internal.SmallBasicApplication
Imports System.Windows

Namespace WinForms
    <SmallBasicType>
    Public NotInheritable Class ListBox

        Private Shared Function GetListBox(formName As String, ListBoxName As String) As Wpf.ListBox
            Dim c = Control.GetControl(formName, ListBoxName)
            Dim t = TryCast(c, Wpf.ListBox)
            If t Is Nothing Then
                Throw New ArgumentException($"{ListBoxName} is not a name of a ListBox.")
            End If
            Return t
        End Function

        <ExProperty>
        Public Shared Function GetItemsCount(formName As Primitive, ListBoxName As Primitive) As Primitive
            App.Invoke(Sub() GetItemsCount = GetListBox(formName, ListBoxName).Items.Count)
        End Function

        <ExProperty>
        Public Shared Function GetSelectedItem(formName As Primitive, ListBoxName As Primitive) As Primitive
            App.Invoke(Sub() GetSelectedItem = GetListBox(formName, ListBoxName).SelectedItem)
        End Function

        <ExProperty>
        Public Shared Sub SetSelectedItem(formName As Primitive, ListBoxName As Primitive, item As Primitive)
            App.Invoke(Sub() GetListBox(formName, ListBoxName).SelectedItem = item)
        End Sub


        <ExProperty>
        Public Shared Function GetSelectedIndex(formName As Primitive, ListBoxName As Primitive) As Primitive
            App.Invoke(Sub() GetSelectedIndex = GetListBox(formName, ListBoxName).SelectedIndex + 1)
        End Function

        <ExProperty>
        Public Shared Sub SetSelectedIndex(formName As Primitive, ListBoxName As Primitive, index As Primitive)
            App.Invoke(
                Sub()
                    Dim lst = GetListBox(formName, ListBoxName)
                    If index < 1 OrElse index > lst.Items.Count Then
                        index = -1
                    Else
                        index -= 1
                    End If
                    lst.SelectedIndex = index
                End Sub)
        End Sub

        <ExMethod>
        Public Shared Function GetItemAt(formName As Primitive, ListBoxName As Primitive, index As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Dim lst = GetListBox(formName, ListBoxName)
                    If index < 1 OrElse index > lst.Items.Count Then
                        GetItemAt = ""
                    Else
                        GetItemAt = lst.Items(index - 1)
                    End If
                End Sub)
        End Function

        <ExMethod>
        Public Shared Sub SetItemAt(formName As Primitive, ListBoxName As Primitive, index As Primitive, item As Primitive)
            App.Invoke(
                Sub()
                    Dim lst = GetListBox(formName, ListBoxName)
                    If index > 0 AndAlso index <= lst.Items.Count Then
                        lst.Items(index - 1) = item
                    End If
                End Sub)
        End Sub

        <ExMethod>
        Public Shared Function AddItem(formName As Primitive, ListBoxName As Primitive, item As Primitive) As Primitive
            App.Invoke(Sub() AddItem = GetListBox(formName, ListBoxName).Items.Add(item))
        End Function

        <ExMethod>
        Public Shared Sub RemoveItem(formName As Primitive, ListBoxName As Primitive, item As Primitive)
            App.Invoke(
                Sub()
                    Dim lst = GetListBox(formName, ListBoxName)
                    Dim i = lst.Items.IndexOf(item)
                    If i > -1 Then lst.Items.RemoveAt(i)
                End Sub)
        End Sub

        <ExMethod>
        Public Shared Sub RemoveItemAt(formName As Primitive, ListBoxName As Primitive, index As Primitive)
            App.Invoke(
                Sub()
                    Dim lst = GetListBox(formName, ListBoxName)
                    If index > 0 AndAlso index <= lst.Items.Count Then
                        lst.Items.RemoveAt(index - 1)
                    End If
                End Sub)
        End Sub

        <ExMethod>
        Public Shared Function ContainsItem(formName As Primitive, ListBoxName As Primitive, item As Primitive) As Primitive
            App.Invoke(Sub() ContainsItem = GetListBox(formName, ListBoxName).Items.Contains(item))
        End Function

        <ExMethod>
        Public Shared Function FindItem(formName As Primitive, ListBoxName As Primitive, item As Primitive) As Primitive
            App.Invoke(Sub() FindItem = 1 + GetListBox(formName, ListBoxName).Items.IndexOf(item))
        End Function

        Public Shared Custom Event OnSelection As SmallBasicCallback
            AddHandler(handler As SmallBasicCallback)
                Dim VisualElement = GetListBox([Event].SenderForm, [Event].SenderControl)
                AddHandler VisualElement.SelectionChanged, Sub(Sender As Wpf.Control, e As RoutedEventArgs) [Event].EventsHandler(Sender, e, handler)
            End AddHandler

            RemoveHandler(handler As SmallBasicCallback)
            End RemoveHandler

            RaiseEvent()
            End RaiseEvent
        End Event

    End Class
End Namespace