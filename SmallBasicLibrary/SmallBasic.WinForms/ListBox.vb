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
                Throw New Exception($"{ListBoxName} is not a name of a ListBox.")
            End If
            Return t
        End Function

        <ExProperty>
        Public Shared Function GetItemsCount(formName As Primitive, ListBoxName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetItemsCount = GetListBox(formName, ListBoxName).Items.Count
                    Catch ex As Exception
                        Control.ShowErrorMesssage(formName, ListBoxName, "ItemsCount", ex.Message)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Function GetSelectedItem(formName As Primitive, ListBoxName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetSelectedItem = GetListBox(formName, ListBoxName).SelectedItem
                    Catch ex As Exception
                        Control.ShowErrorMesssage(formName, ListBoxName, "SelectedItem", ex.Message)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetSelectedItem(formName As Primitive, ListBoxName As Primitive, item As Primitive)
            App.Invoke(
                Sub()
                    Try
                        GetListBox(formName, ListBoxName).SelectedItem = item
                    Catch ex As Exception
                        Control.ShowErrorMesssage(formName, ListBoxName, "SelectedItem", item, ex.Message)
                    End Try
                End Sub)
        End Sub


        <ExProperty>
        Public Shared Function GetSelectedIndex(formName As Primitive, ListBoxName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetSelectedIndex = GetListBox(formName, ListBoxName).SelectedIndex + 1
                    Catch ex As Exception
                        Control.ShowErrorMesssage(formName, ListBoxName, "SelectedIndex", ex.Message)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetSelectedIndex(formName As Primitive, ListBoxName As Primitive, index As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Dim lst = GetListBox(formName, ListBoxName)
                        If index < 1 OrElse index > lst.Items.Count Then
                            index = -1
                        Else
                            index -= 1
                        End If
                        lst.SelectedIndex = index
                    Catch ex As Exception
                        Control.ShowErrorMesssage(formName, ListBoxName, "SelectedIndex", index, ex.Message)
                    End Try
                End Sub)
        End Sub

        <ExMethod>
        Public Shared Function GetItemAt(formName As Primitive, ListBoxName As Primitive, index As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        Dim lst = GetListBox(formName, ListBoxName)
                        If index < 1 OrElse index > lst.Items.Count Then
                            GetItemAt = ""
                        Else
                            GetItemAt = lst.Items(index - 1)
                        End If
                    Catch ex As Exception
                        Control.ShowSubError(formName, ListBoxName, "GetItemAt", ex.Message)
                    End Try
                End Sub)
        End Function

        <ExMethod>
        Public Shared Sub SetItemAt(formName As Primitive, ListBoxName As Primitive, index As Primitive, item As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Dim lst = GetListBox(formName, ListBoxName)
                        If index > 0 AndAlso index <= lst.Items.Count Then
                            lst.Items(index - 1) = item
                        End If
                    Catch ex As Exception
                        Control.ShowSubError(formName, ListBoxName, "SetItemAt", ex.Message)
                    End Try
                End Sub)
        End Sub

        <ExMethod>
        Public Shared Function AddItem(formName As Primitive, ListBoxName As Primitive, item As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        Dim lst = GetListBox(formName, ListBoxName)
                        If item.IsArray Then
                            For Each value In item._arrayMap.Values
                                AddItem = lst.Items.Add(value)
                            Next
                        Else
                            AddItem = lst.Items.Add(item)
                        End If
                    Catch ex As Exception
                        Control.ShowSubError(formName, ListBoxName, "AddItem", ex.Message)
                    End Try
                End Sub)
        End Function

        <ExMethod>
        Public Shared Sub AddItemAt(formName As Primitive, ListBoxName As Primitive, item As Primitive, index As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Dim lst = GetListBox(formName, ListBoxName)
                        If item.IsArray Then
                            Dim items = item._arrayMap.Values
                            For i = 0 To items.Count - 1
                                lst.Items.Insert(index + 1, items(i))
                            Next
                        Else
                            lst.Items.Insert(index, item)
                        End If
                    Catch ex As Exception
                        Control.ShowSubError(formName, ListBoxName, "AddItemAt", ex.Message)
                    End Try
                End Sub)
        End Sub

        <ExMethod>
        Public Shared Sub RemoveItem(formName As Primitive, ListBoxName As Primitive, item As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Dim lst = GetListBox(formName, ListBoxName)
                        Dim i = lst.Items.IndexOf(item)
                        If i > -1 Then lst.Items.RemoveAt(i)
                    Catch ex As Exception
                        Control.ShowSubError(formName, ListBoxName, "RenoveItem", ex.Message)
                    End Try
                End Sub)
        End Sub

        <ExMethod>
        Public Shared Sub RemoveItemAt(formName As Primitive, ListBoxName As Primitive, index As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Dim lst = GetListBox(formName, ListBoxName)
                        If index > 0 AndAlso index <= lst.Items.Count Then
                            lst.Items.RemoveAt(index - 1)
                        End If
                    Catch ex As Exception
                        Control.ShowSubError(formName, ListBoxName, "RenoveItemAt", ex.Message)
                    End Try
                End Sub)
        End Sub

        <ExMethod>
        Public Shared Function ContainsItem(formName As Primitive, ListBoxName As Primitive, item As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        ContainsItem = GetListBox(formName, ListBoxName).Items.Contains(item)
                    Catch ex As Exception
                        Control.ShowSubError(formName, ListBoxName, "ContainsItem", ex.Message)
                    End Try
                End Sub)
        End Function

        <ExMethod>
        Public Shared Function FindItem(formName As Primitive, ListBoxName As Primitive, item As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        FindItem = 1 + GetListBox(formName, ListBoxName).Items.IndexOf(item)
                    Catch ex As Exception
                        Control.ShowSubError(formName, ListBoxName, "FindItem", ex.Message)
                    End Try
                End Sub)
        End Function

        Public Shared Custom Event OnSelection As SmallBasicCallback
            AddHandler(handler As SmallBasicCallback)
                Try
                    Dim VisualElement = GetListBox([Event].SenderForm, [Event].SenderControl)
                    AddHandler VisualElement.SelectionChanged, Sub(Sender As Wpf.Control, e As RoutedEventArgs) [Event].EventsHandler(Sender, e, handler)
                Catch ex As Exception
                    [Event].ShowErrorMessage(NameOf(OnSelection), ex.Message)
                End Try

            End AddHandler

            RemoveHandler(handler As SmallBasicCallback)
            End RemoveHandler

            RaiseEvent()
            End RaiseEvent
        End Event

    End Class
End Namespace