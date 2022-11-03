
Imports Microsoft.SmallVisualBasic.Library
Imports Wpf = System.Windows.Controls
Imports App = Microsoft.SmallVisualBasic.Library.Internal.SmallBasicApplication
Imports System.Windows

Namespace WinForms
    <SmallVisualBasicType>
    <HideFromIntellisense>
    Public NotInheritable Class ComboBox

        Private Shared Function GetComboBox(comboBoxName As String) As Wpf.ComboBox
            Dim c = Control.GetControl(comboBoxName)
            Dim t = TryCast(c, Wpf.ComboBox)
            If t Is Nothing Then
                Throw New Exception($"{comboBoxName} is not a name of a ComboBox.")
            End If
            Return t
        End Function


        ''' <summary>
        ''' Gets the count of the items in the ComboBox
        ''' </summary>
        <ReturnValueType(VariableType.Double)>
        <ExProperty>
        Public Shared Function GetItemsCount(comboBoxName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetItemsCount = GetComboBox(comboBoxName).Items.Count
                    Catch ex As Exception
                        Control.ReportError(comboBoxName, "ItemsCount", ex)
                    End Try
                End Sub)
        End Function

        ''' <summary>
        ''' Gets an array containing the items of the ComboBox
        ''' </summary>
        <ReturnValueType(VariableType.Array)>
        <ExProperty>
        Public Shared Function GetItems(comboBoxName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        Dim map = New Dictionary(Of Primitive, Primitive)
                        Dim num = 1
                        For Each item In GetComboBox(comboBoxName).Items
                            map(num) = CStr(item)
                            num += 1
                        Next
                        GetItems = Primitive.ConvertFromMap(map)

                    Catch ex As Exception
                        Control.ReportError(comboBoxName, "ItemsCount", ex)
                    End Try
                End Sub)
        End Function

        ''' <summary>
        ''' Gets or sets the item that is curruntly selected in the ComboBox
        ''' </summary>
        ''' <remarks>This property returns empty string if there is no item selected.
        ''' But there can ba a selected item that displays an empty string!
        ''' So, use the SelectedIndex property if you want to distinguish between the two cases.
        ''' If you set the selectedItem to a value that doesn't existed in the list, 
        ''' no item will be selected.
        ''' </remarks>
        <ReturnValueType(VariableType.String)>
        <ExProperty>
        Public Shared Function GetSelectedItem(comboBoxName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        Dim item As String = GetComboBox(comboBoxName).SelectedItem
                        GetSelectedItem = item
                    Catch ex As Exception
                        Control.ReportError(comboBoxName, "SelectedItem", ex)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetSelectedItem(comboBoxName As Primitive, item As Primitive)
            App.Invoke(
                Sub()
                    Try
                        GetComboBox(comboBoxName).SelectedItem = CStr(item)
                    Catch ex As Exception
                        Control.ReportError(comboBoxName, "SelectedItem", item, ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Gets or sets the index of the selected item in the ComboBox. 
        ''' Zero indicates that no item is selected.
        ''' </summary>
        <ReturnValueType(VariableType.Double)>
        <ExProperty>
        Public Shared Function GetSelectedIndex(comboBoxName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetSelectedIndex = GetComboBox(comboBoxName).SelectedIndex + 1
                    Catch ex As Exception
                        Control.ReportError(comboBoxName, "SelectedIndex", ex)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetSelectedIndex(comboBoxName As Primitive, index As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Dim lst = GetComboBox(comboBoxName)
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
                        Control.ReportError(comboBoxName, "SelectedIndex", index, ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Set this property to True to allow the user to write in the textbox of the ComboBox.
        ''' </summary>
        <ReturnValueType(VariableType.Boolean)>
        <ExProperty>
        Public Shared Function GetAllowEdit(comboBoxName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetAllowEdit = GetComboBox(comboBoxName).IsEditable
                    Catch ex As Exception
                        Control.ReportError(comboBoxName, "AllowEdit", ex)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetAllowEdit(comboBoxName As Primitive, item As Primitive)
            App.Invoke(
                Sub()
                    Try
                        GetComboBox(comboBoxName).IsEditable = CBool(item)
                    Catch ex As Exception
                        Control.ReportError(comboBoxName, "AllowEdit", item, ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Gets or sets the text that is displayed by the textbox of the ComboBox
        ''' </summary>
        <ReturnValueType(VariableType.String)>
        <ExProperty>
        Public Shared Function GetText(comboBoxName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetText = GetComboBox(comboBoxName).Text
                    Catch ex As Exception
                        Control.ReportError(comboBoxName, "Text", ex)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetText(textBoxName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        GetComboBox(textBoxName).Text = value
                    Catch ex As Exception
                        Control.ReportPropertyError(textBoxName, "Text", value, ex)
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
        Public Shared Function GetItemAt(comboBoxName As Primitive, index As Primitive) As Primitive
            App.Invoke(
                Sub()

                    Try
                        Dim lst = GetComboBox(comboBoxName)
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
                        Control.ReportSubError(comboBoxName, "GetItemAt", ex)
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
        Public Shared Sub SetItemAt(comboBoxName As Primitive, index As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        If Not index.IsNumber Then Return

                        Dim lst = GetComboBox(comboBoxName)
                        Dim i As Integer = index
                        If i > 0 AndAlso i <= lst.Items.Count Then
                            lst.Items(i - 1) = value
                        End If
                    Catch ex As Exception
                        Control.ReportSubError(comboBoxName, "SetItemAt", ex)
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
        Public Shared Function AddItem(comboBoxName As Primitive, value As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        Dim lst = GetComboBox(comboBoxName)
                        If value.IsArray Then
                            For Each value In value._arrayMap.Values
                                AddItem = lst.Items.Add(CStr(value))
                            Next
                        Else
                            AddItem = lst.Items.Add(CStr(value))
                        End If

                    Catch ex As Exception
                        Control.ReportSubError(comboBoxName, "AddItem", ex)
                    End Try
                End Sub)
        End Function

        ''' <summary>
        ''' Adds the given item to the list at the given index.
        ''' </summary>
        ''' <param name="value">the item you want to add to the list. You can send an array to add all its items</param>
        ''' <param name="index">the index to want to add the item at</param>
        <ExMethod>
        Public Shared Sub AddItemAt(comboBoxName As Primitive, value As Primitive, index As Primitive)
            App.Invoke(
                Sub()
                    Try
                        If Not index.IsNumber Then Return
                        Dim i As Integer = CInt(index) - 1
                        Dim lst = GetComboBox(comboBoxName)
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
                        Control.ReportSubError(comboBoxName, "AddItemAt", ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Searches for the given value in the list, and reomves the first item if founds.
        ''' </summary>
        ''' <param name="value">the item you want to remove. You can send an array to remove all its items</param>
        <ExMethod>
        Public Shared Sub RemoveItem(comboBoxName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Dim lst = GetComboBox(comboBoxName)
                        If value.IsArray Then
                            For Each value In value._arrayMap.Values
                                RemoveItem(value, lst)
                            Next
                        Else
                            RemoveItem(value, lst)
                        End If

                    Catch ex As Exception
                        Control.ReportSubError(comboBoxName, "RenoveItem", ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Reomves all the items from the listbox
        ''' </summary>
        <ExMethod>
        Public Shared Sub RemoveAllItems(comboBoxName As Primitive)
            App.Invoke(
                Sub()
                    Try
                        GetComboBox(comboBoxName).Items.Clear()
                    Catch ex As Exception
                        Control.ReportSubError(comboBoxName, "RenoveItem", ex)
                    End Try
                End Sub)
        End Sub

        Private Shared Sub RemoveItem(item As String, lst As Wpf.ComboBox)
            Dim i = lst.Items.IndexOf(item)
            If i > -1 Then lst.Items.RemoveAt(i)
        End Sub

        ''' <summary>
        ''' Removes the list item that exists at the given index
        ''' </summary>
        ''' <param name="index">The index of the item you want to remove</param>
        <ExMethod>
        Public Shared Sub RemoveItemAt(comboBoxName As Primitive, index As Primitive)
            App.Invoke(
                Sub()
                    Try
                        If Not index.IsNumber Then Return

                        Dim lst = GetComboBox(comboBoxName)
                        Dim i As Integer = index
                        If i > 0 AndAlso i <= lst.Items.Count Then
                            lst.Items.RemoveAt(i - 1)
                        End If

                    Catch ex As Exception
                        Control.ReportSubError(comboBoxName, "RenoveItemAt", ex)
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
        Public Shared Function ContainsItem(comboBoxName As Primitive, value As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        ContainsItem = GetComboBox(comboBoxName).Items.Contains(CStr(value))
                    Catch ex As Exception
                        Control.ReportSubError(comboBoxName, "ContainsItem", ex)
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
        Public Shared Function FindItem(comboBoxName As Primitive, value As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        FindItem = 1 + GetComboBox(comboBoxName).Items.IndexOf(CStr(value))
                    Catch ex As Exception
                        Control.ReportSubError(comboBoxName, "FindItem", ex)
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
        Public Shared Function FindItemAt(comboBoxName As Primitive, value As Primitive, startIndex As Primitive, endIndex As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        Dim lst = GetComboBox(comboBoxName)
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
                        Control.ReportSubError(comboBoxName, "FindItemAt", ex)
                    End Try
                End Sub)
        End Function

        ''' <summary>
        ''' This event is fired when the selected item in the list is changed.
        ''' </summary>
        Public Shared Custom Event OnSelection As SmallBasicCallback
            AddHandler(handler As SmallBasicCallback)
                Try
                    Dim _sender = GetComboBox([Event].SenderControl)
                    AddHandler _sender.SelectionChanged, Sub(Sender As Wpf.Control, e As RoutedEventArgs) [Event].EventsHandler(CType(Sender, FrameworkElement), e, handler)
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