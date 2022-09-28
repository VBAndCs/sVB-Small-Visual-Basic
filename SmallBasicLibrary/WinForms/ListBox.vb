
Imports Microsoft.SmallVisualBasic.Library
Imports Wpf = System.Windows.Controls
Imports App = Microsoft.SmallVisualBasic.Library.Internal.SmallBasicApplication
Imports System.Windows

Namespace WinForms
    <SmallBasicType>
    <HideFromIntellisense>
    Public NotInheritable Class ListBox

        Private Shared Function GetListBox(ListBoxName As String) As Wpf.ListBox
            Dim c = Control.GetControl(ListBoxName)
            Dim t = TryCast(c, Wpf.ListBox)
            If t Is Nothing Then
                Throw New Exception($"{ListBoxName} is not a name of a ListBox.")
            End If
            Return t
        End Function


        ''' <summary>
        ''' Gets the count of the items in the ListBox
        ''' </summary>
        <ReturnValueType(VariableType.Double)>
        <ExProperty>
        Public Shared Function GetItemsCount(ListBoxName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetItemsCount = GetListBox(ListBoxName).Items.Count
                    Catch ex As Exception
                        Control.ShowErrorMesssage(ListBoxName, "ItemsCount", ex)
                    End Try
                End Sub)
        End Function

        ''' <summary>
        ''' Gets an array containing the items of the ListBox
        ''' </summary>
        <ReturnValueType(VariableType.Array)>
        <ExProperty>
        Public Shared Function GetItems(ListBoxName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        Dim map = New Dictionary(Of Primitive, Primitive)
                        Dim num = 1
                        For Each item In GetListBox(ListBoxName).Items
                            map(num) = CStr(item)
                            num += 1
                        Next
                        GetItems = Primitive.ConvertFromMap(map)

                    Catch ex As Exception
                        Control.ShowErrorMesssage(ListBoxName, "ItemsCount", ex)
                    End Try
                End Sub)
        End Function

        ''' <summary>
        ''' Gets or sets the item that is curruntly selected in the ListBox
        ''' </summary>
        ''' <remarks>This property returns empty string if there is no item selected.
        ''' But some there can ba a selected item that displays an empty string!
        ''' So, use the SelectedIndex property if you want to distinguish between the two cases.
        ''' If you set the selectedItem to a value that doesn't existed in the list, 
        ''' no item will be selected.
        ''' </remarks>
        <ReturnValueType(VariableType.String)>
        <ExProperty>
        Public Shared Function GetSelectedItem(ListBoxName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        Dim item As String = GetListBox(ListBoxName).SelectedItem
                        GetSelectedItem = item
                    Catch ex As Exception
                        Control.ShowErrorMesssage(ListBoxName, "SelectedItem", ex)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetSelectedItem(ListBoxName As Primitive, item As Primitive)
            App.Invoke(
                Sub()
                    Try
                        GetListBox(ListBoxName).SelectedItem = CStr(item)
                    Catch ex As Exception
                        Control.ShowErrorMesssage(ListBoxName, "SelectedItem", item, ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Gets or sets the index of the selected item in the ListBox. 
        ''' Zero indicates that no item is selected.
        ''' </summary>
        <ReturnValueType(VariableType.Double)>
        <ExProperty>
        Public Shared Function GetSelectedIndex(ListBoxName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetSelectedIndex = GetListBox(ListBoxName).SelectedIndex + 1
                    Catch ex As Exception
                        Control.ShowErrorMesssage(ListBoxName, "SelectedIndex", ex)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetSelectedIndex(ListBoxName As Primitive, index As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Dim lst = GetListBox(ListBoxName)
                        Dim i As Integer
                        If Not index.IsNumber Then
                            i = -1
                        Else
                            i = index
                        End If

                        If i < 1 OrElse i > lst.Items.Count Then
                            i = -1
                        Else
                            i -= 1
                        End If

                        lst.SelectedIndex = i

                    Catch ex As Exception
                        Control.ShowErrorMesssage(ListBoxName, "SelectedIndex", index, ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Returns the item that exists in the given index in the list.
        ''' </summary>
        ''' <param name="index">
        ''' The index of the item. It should be greater than zero and not exceed the count of the items,
        ''' otherwise, this method will return an empty string.
        ''' </param>
        ''' <returns>the value of the item</returns>
        <ReturnValueType(VariableType.String)>
        <ExMethod>
        Public Shared Function GetItemAt(ListBoxName As Primitive, index As Primitive) As Primitive
            App.Invoke(
                Sub()

                    Try
                        Dim lst = GetListBox(ListBoxName)
                        If Not index.IsNumber Then
                            GetItemAt = ""
                        Else
                            Dim i As Integer = index
                            If i < 1 OrElse i > lst.Items.Count Then
                                GetItemAt = ""
                            Else
                                GetItemAt = lst.Items(i - 1).ToString()
                            End If
                        End If
                    Catch ex As Exception
                        Control.ShowSubError(ListBoxName, "GetItemAt", ex)
                    End Try
                End Sub)
        End Function

        ''' <summary>
        ''' Sets the value of the item that exists in the given index in the list.
        ''' </summary>
        ''' <param name="index">
        ''' The index of the item. It should be greater than zero and not exceed the count of the items,
        ''' otherwise, this method will do nothing.
        ''' </param>
        ''' <param name="value">the value to set to the item</param>
        <ExMethod>
        Public Shared Sub SetItemAt(ListBoxName As Primitive, index As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        If Not index.IsNumber Then Return

                        Dim lst = GetListBox(ListBoxName)
                        Dim i As Integer = index
                        If i > 0 AndAlso i <= lst.Items.Count Then
                            lst.Items(i - 1) = value
                        End If
                    Catch ex As Exception
                        Control.ShowSubError(ListBoxName, "SetItemAt", ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Adds an item to the end of the list.
        ''' </summary>
        ''' <param name="value">The item you want to add to the list. You can send an array to add all its items</param>
        ''' <returns>the index of the newly added item</returns>
        <ReturnValueType(VariableType.Double)>
        <ExMethod>
        Public Shared Function AddItem(ListBoxName As Primitive, value As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        Dim lst = GetListBox(ListBoxName)
                        If value.IsArray Then
                            For Each value In value._arrayMap.Values
                                AddItem = lst.Items.Add(CStr(value))
                            Next
                        Else
                            AddItem = lst.Items.Add(CStr(value))
                        End If

                    Catch ex As Exception
                        Control.ShowSubError(ListBoxName, "AddItem", ex)
                    End Try
                End Sub)
        End Function

        ''' <summary>
        ''' Adds the given item to the list at the given index.
        ''' </summary>
        ''' <param name="value">the item you want to add to the list. You can send an array to add all its items</param>
        ''' <param name="index">the index to want to add the item at</param>
        <ExMethod>
        Public Shared Sub AddItemAt(ListBoxName As Primitive, value As Primitive, index As Primitive)
            App.Invoke(
                Sub()
                    Try
                        If Not index.IsNumber Then Return
                        Dim i As Integer = CInt(index) - 1
                        Dim lst = GetListBox(ListBoxName)
                        If i < 0 OrElse i > lst.Items.Count Then Return

                        If value.IsArray Then
                            Dim items = value._arrayMap.Values
                            For n = items.Count - 1 To 0 Step -1
                                lst.Items.Insert(i, CStr(items(n)))
                            Next
                        Else
                            lst.Items.Insert(i, value)
                        End If

                    Catch ex As Exception
                        Control.ShowSubError(ListBoxName, "AddItemAt", ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Searches for the given value in the list, and reomves the first item if founds.
        ''' </summary>
        ''' <param name="value">the item you want to remove. You can send an array to remove all its items</param>
        <ExMethod>
        Public Shared Sub RemoveItem(ListBoxName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Dim lst = GetListBox(ListBoxName)
                        If value.IsArray Then
                            For Each value In value._arrayMap.Values
                                RemoveItem(value, lst)
                            Next
                        Else
                            RemoveItem(value, lst)
                        End If

                    Catch ex As Exception
                        Control.ShowSubError(ListBoxName, "RenoveItem", ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Reomves all the items from the listbox
        ''' </summary>
        <ExMethod>
        Public Shared Sub RemoveAllItems(ListBoxName As Primitive)
            App.Invoke(
                Sub()
                    Try
                        GetListBox(ListBoxName).Items.Clear()
                    Catch ex As Exception
                        Control.ShowSubError(ListBoxName, "RenoveItem", ex)
                    End Try
                End Sub)
        End Sub

        Private Shared Sub RemoveItem(item As String, lst As Wpf.ListBox)
            Dim i = lst.Items.IndexOf(item)
            If i > -1 Then lst.Items.RemoveAt(i)
        End Sub

        ''' <summary>
        ''' Removes the list item that exists at the given index
        ''' </summary>
        ''' <param name="index">The index of the item you want to remove</param>
        <ExMethod>
        Public Shared Sub RemoveItemAt(ListBoxName As Primitive, index As Primitive)
            App.Invoke(
                Sub()
                    Try
                        If Not index.IsNumber Then Return

                        Dim lst = GetListBox(ListBoxName)
                        Dim i As Integer = index
                        If i > 0 AndAlso i <= lst.Items.Count Then
                            lst.Items.RemoveAt(i - 1)
                        End If

                    Catch ex As Exception
                        Control.ShowSubError(ListBoxName, "RenoveItemAt", ex)
                    End Try
                End Sub)
        End Sub


        ''' <summary>
        ''' Returns True if the list contains the given item.
        ''' </summary>
        ''' <param name="value">the value of the item you ant to check it exists in the list</param>
        ''' <returns>True if the item found or False otherwise</returns>
        <ReturnValueType(VariableType.Boolean)>
        <ExMethod>
        Public Shared Function ContainsItem(ListBoxName As Primitive, value As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        ContainsItem = GetListBox(ListBoxName).Items.Contains(CStr(value))
                    Catch ex As Exception
                        Control.ShowSubError(ListBoxName, "ContainsItem", ex)
                    End Try
                End Sub)
        End Function

        ''' <summary>
        ''' Returns the index of the given item if it esists in the list, otherwise retruns 0.
        ''' </summary>
        ''' <param name="value">The item you want to find</param>
        ''' <returns>the index of the item if found, or 0 otherwise</returns>
        <ReturnValueType(VariableType.Double)>
        <ExMethod>
        Public Shared Function FindItem(ListBoxName As Primitive, value As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        FindItem = 1 + GetListBox(ListBoxName).Items.IndexOf(CStr(value))
                    Catch ex As Exception
                        Control.ShowSubError(ListBoxName, "FindItem", ex)
                    End Try
                End Sub)
        End Function

        ''' <summary>
        ''' Returns the index of the given item if it esists in the list in the given index range, otherwise retruns 0.
        ''' </summary>
        ''' <param name="value">The item you want to fin</param>
        ''' <param name="startIndex">the array index the seach starts at</param>
        ''' <param name="endIndex">the array index the seach ends at</param>
        ''' <returns>the index of the item if found, or 0 otherwise</returns>
        <ExMethod>
        <ReturnValueType(VariableType.Double)>
        Public Shared Function FindItemAt(ListBoxName As Primitive, value As Primitive, startIndex As Primitive, endIndex As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        Dim lst = GetListBox(ListBoxName)
                        FindItemAt = 0

                        If Not startIndex.IsNumber OrElse Not endIndex.IsNumber Then
                            Return
                        End If

                        Dim items = lst.Items
                        Dim st As Integer = System.Math.Max(0, startIndex - 1)
                        Dim en As Integer = System.Math.Min(endIndex - 1, items.Count - 1)
                        Dim str = CStr(value)

                        For i = st To en Step If(st > en, -1, 1)
                            If CStr(items(i)) = str Then
                                FindItemAt = i + 1
                                Return
                            End If
                        Next

                    Catch ex As Exception
                        Control.ShowSubError(ListBoxName, "FindItemAt", ex)
                    End Try
                End Sub)
        End Function

        ''' <summary>
        ''' This event is fired when the selected item in the list is changed.
        ''' </summary>
        Public Shared Custom Event OnSelection As SmallBasicCallback
            AddHandler(handler As SmallBasicCallback)
                Try
                    Dim VisualElement = GetListBox([Event].SenderControl)
                    AddHandler VisualElement.SelectionChanged, Sub(Sender As Wpf.Control, e As RoutedEventArgs) [Event].EventsHandler(CType(Sender, FrameworkElement), e, handler)
                Catch ex As Exception
                    [Event].ShowErrorMessage(NameOf(OnSelection), ex)
                End Try

            End AddHandler

            RemoveHandler(handler As SmallBasicCallback)
            End RemoveHandler

            RaiseEvent()
            End RaiseEvent
        End Event

    End Class
End Namespace