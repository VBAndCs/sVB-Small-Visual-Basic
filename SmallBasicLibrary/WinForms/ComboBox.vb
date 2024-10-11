
Imports Microsoft.SmallVisualBasic.Library
Imports Wpf = System.Windows.Controls
Imports App = Microsoft.SmallVisualBasic.Library.Internal.SmallBasicApplication
Imports System.Windows
Imports System.Windows.Media

Namespace WinForms
    ''' <summary>
    ''' Represents a ComboBox control, which is composed of a textbox and a dropdown listbox.
    ''' You can use the form designer to add a combo box to the form by dragging it from the toolbox.
    ''' It is also possible to use the Form.AddComboBox method to create a new combo box and add it to the form at runtime.
    ''' </summary>
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
            Return ListBase.GetItemsCount(comboBoxName)
        End Function

        ''' <summary>
        ''' Gets an array containing the items of the ComboBox
        ''' </summary>
        <ReturnValueType(VariableType.Array)>
        <ExProperty>
        Public Shared Function GetItems(comboBoxName As Primitive) As Primitive
            Return ListBase.GetItems(comboBoxName)
        End Function

        ''' <summary>
        ''' Gets or sets the item that is currently selected in the ComboBox
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
            Return ListBase.GetSelectedItem(comboBoxName)
        End Function

        <ExProperty>
        Public Shared Sub SetSelectedItem(comboBoxName As Primitive, item As Primitive)
            ListBase.SetSelectedItem(comboBoxName, item)
        End Sub

        ''' <summary>
        ''' Gets or sets the index of the selected item in the ComboBox. 
        ''' Zero indicates that no item is selected.
        ''' </summary>
        <ReturnValueType(VariableType.Double)>
        <ExProperty>
        Public Shared Function GetSelectedIndex(comboBoxName As Primitive) As Primitive
            Return ListBase.GetSelectedIndex(comboBoxName)
        End Function

        <ExProperty>
        Public Shared Sub SetSelectedIndex(comboBoxName As Primitive, index As Primitive)
            ListBase.SetSelectedIndex(comboBoxName, index)
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
                        Control.ReportPropertyError(comboBoxName, "AllowEdit", item, ex)
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
                        Dim cmb = GetComboBox(comboBoxName)
                        GetText = New Primitive(cmb.Text)
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
                        GetComboBox(textBoxName).Text = value.AsString()
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
            Return ListBase.GetItemAt(comboBoxName, index)
        End Function

        ''' <summary>
        ''' Sets the value of the item that exists in the given index in the list.
        ''' </summary>
        ''' <param name="index">The index of the item. It should be greater than zero and not exceed the count of the items</param>
        ''' <param name="value">the value to set to the item</param>
        ''' <returns>True if the item is modified, otherwise False.</returns>
        <ExMethod>
        <ReturnValueType(VariableType.Boolean)>
        Public Shared Function SetItemAt(comboBoxName As Primitive, index As Primitive, value As Primitive) As Primitive
            Return ListBase.SetItemAt(comboBoxName, index, value)
        End Function

        ''' <summary>
        ''' Adds an item to the end of the list.
        ''' </summary>
        ''' <param name="value">The item you want to add to the list. You can send an array to add all its items</param>
        ''' <returns>the index of the newly added item, or 0 if the operation failed.</returns>
        <ReturnValueType(VariableType.Double)>
        <ExMethod>
        Public Shared Function AddItem(comboBoxName As Primitive, value As Primitive) As Primitive
            Return ListBase.AddItem(comboBoxName, value)
        End Function

        ''' <summary>
        ''' Adds the given item to the list at the given index.
        ''' </summary>
        ''' <param name="value">the item you want to add to the list. You can send an array to add all its items</param>
        ''' <param name="index">The index you want to add the item at. The value of this index must be greater that 0 and less that list items count + 1, otherwise the item will not be added.</param>
        ''' <returns>True is the item is successfully added at the given index, otherwise False.</returns>
        <ExMethod>
        <ReturnValueType(VariableType.Boolean)>
        Public Shared Function AddItemAt(comboBoxName As Primitive, value As Primitive, index As Primitive) As Primitive
            Return ListBase.AddItemAt(comboBoxName, value, index)
        End Function

        ''' <summary>
        ''' Searches for the given value in the list, and removes the first found item.
        ''' </summary>
        ''' <param name="value">the item you want to remove. You can send an array to remove all its items</param>
        ''' <returns>True is the item is removed, otherwise False.</returns>
        <ExMethod>
        <ReturnValueType(VariableType.Boolean)>
        Public Shared Function RemoveItem(comboBoxName As Primitive, value As Primitive) As Primitive
            Return ListBase.RemoveItem(comboBoxName, value)
        End Function

        ''' <summary>
        ''' Reomves all the items from the listbox
        ''' </summary>
        <ExMethod>
        Public Shared Sub RemoveAllItems(comboBoxName As Primitive)
            ListBase.RemoveAllItems(comboBoxName)
        End Sub

        ''' <summary>
        ''' Removes the list item that exists at the given index
        ''' </summary>
        ''' <param name="index">The index of the item you want to remove</param>
        ''' <returns>True if the item id removd, otherwise False.</returns>
        <ExMethod>
        <ReturnValueType(VariableType.Boolean)>
        Public Shared Function RemoveItemAt(comboBoxName As Primitive, index As Primitive) As Primitive
            Return ListBase.RemoveItemAt(comboBoxName, index)
        End Function


        ''' <summary>
        ''' Checkes whether or not the given item exists in the list.
        ''' </summary>
        ''' <param name="value">The value of the item to search for.</param>
        ''' <returns>True if the item found, or False otherwise.</returns>
        <ReturnValueType(VariableType.Boolean)>
        <ExMethod>
        Public Shared Function ContainsItem(comboBoxName As Primitive, value As Primitive) As Primitive
            Return ListBase.ContainsItem(comboBoxName, value)
        End Function

        ''' <summary>
        ''' Returns the index of the given item if it esists in the list, otherwise retruns 0.
        ''' </summary>
        ''' <param name="value">The item you want to find</param>
        ''' <returns>the index of the item if found, or 0 otherwise</returns>
        <ReturnValueType(VariableType.Double)>
        <ExMethod>
        Public Shared Function FindItem(comboBoxName As Primitive, value As Primitive) As Primitive
            Return ListBase.FindItem(comboBoxName, value)
        End Function

        ''' <summary>
        ''' Returns the index of the given item if it esists in the list in the given index range, otherwise retruns 0.
        ''' </summary>
        ''' <param name="value">The item you want to find.</param>
        ''' <param name="startIndex">The array index the search starts at.</param>
        ''' <param name="endIndex">The array index the search ends at.If endIndex is less than startIndex, the search direction will be reversed to find the last index of the item.</param>
        ''' <returns>the index of the item if found, or 0 otherwise</returns>
        <ExMethod>
        <ReturnValueType(VariableType.Double)>
        Public Shared Function FindItemAt(comboBoxName As Primitive, value As Primitive, startIndex As Primitive, endIndex As Primitive) As Primitive
            Return ListBase.FindItemAt(comboBoxName, value, startIndex, endIndex)
        End Function

        Friend Shared Sub SetBackColor(cmb As Wpf.ComboBox, color As Media.Color?)
            Dim brush As SolidColorBrush = Nothing
            If color IsNot Nothing Then
                brush = New SolidColorBrush(color.Value)
            End If
            Dim txt = CType(cmb.Template.FindName("PART_EditableTextBox", cmb), Wpf.TextBox)
            txt.Background = brush
        End Sub

        ''' <summary>
        ''' This event is fired when the selected item in the list is changed.
        ''' </summary>
        Public Shared Custom Event OnSelection As SmallVisualBasicCallback
            AddHandler(handler As SmallVisualBasicCallback)
                ListBase.AddOnSelectHandler(handler)
            End AddHandler

            RemoveHandler(handler As SmallVisualBasicCallback)
            End RemoveHandler

            RaiseEvent()
            End RaiseEvent
        End Event

    End Class
End Namespace