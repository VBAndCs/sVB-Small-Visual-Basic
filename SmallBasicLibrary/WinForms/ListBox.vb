
Imports Microsoft.SmallVisualBasic.Library
Imports Wpf = System.Windows.Controls
Imports App = Microsoft.SmallVisualBasic.Library.Internal.SmallBasicApplication
Imports System.Windows

Namespace WinForms
    ''' <summary>
    ''' Represents a ListBox control, which shows a list of items to the user to select one of them.
    ''' You can use the form designer to add a list box to the form by dragging it from the toolbox.
    ''' It is also possible to use the Form.AddListBox method to create a new list box and add it to the form at runtime.
    ''' </summary>
    <SmallVisualBasicType>
    <HideFromIntellisense>
    Public NotInheritable Class ListBox

        Private Shared Function GetListBox(listBoxName As String) As Wpf.ListBox
            Dim c = Control.GetControl(listBoxName)
            Dim lst = TryCast(c, Wpf.ListBox)
            If lst Is Nothing Then
                Throw New Exception($"{listBoxName} is not a name of a ListBox.")
            End If
            Return lst
        End Function


        ''' <summary>
        ''' Gets the count of the items in the ListBox
        ''' </summary>
        <ReturnValueType(VariableType.Double)>
        <ExProperty>
        Public Shared Function GetItemsCount(listBoxName As Primitive) As Primitive
            Return ListBase.GetItemsCount(listBoxName)
        End Function

        ''' <summary>
        ''' Gets an array containing the items of the ListBox
        ''' </summary>
        <ReturnValueType(VariableType.Array)>
        <ExProperty>
        Public Shared Function GetItems(listBoxName As Primitive) As Primitive
            Return ListBase.GetItems(listBoxName)
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
        Public Shared Function GetSelectedItem(listBoxName As Primitive) As Primitive
            Return ListBase.GetSelectedItem(listBoxName)
        End Function

        <ExProperty>
        Public Shared Sub SetSelectedItem(listBoxName As Primitive, item As Primitive)
            ListBase.SetSelectedItem(listBoxName, item)
        End Sub

        ''' <summary>
        ''' Gets or sets the index of the selected item in the ListBox. 
        ''' Zero indicates that no item is selected.
        ''' </summary>
        <ReturnValueType(VariableType.Double)>
        <ExProperty>
        Public Shared Function GetSelectedIndex(listBoxName As Primitive) As Primitive
            Return ListBase.GetSelectedIndex(listBoxName)
        End Function

        <ExProperty>
        Public Shared Sub SetSelectedIndex(listBoxName As Primitive, index As Primitive)
            ListBase.SetSelectedIndex(listBoxName, index)
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
        Public Shared Function GetItemAt(listBoxName As Primitive, index As Primitive) As Primitive
            Return ListBase.GetItemAt(listBoxName, index)
        End Function

        ''' <summary>
        ''' Sets the value of the item that exists in the given index in the list.
        ''' </summary>
        ''' <param name="index">The index of the item. It should be greater than zero and not exceed the count of the items.</param>
        ''' <param name="value">the value to set to the item</param>
        ''' <returns>True if the item is modified, otherwise False.</returns>
        <ExMethod>
        <ReturnValueType(VariableType.Boolean)>
        Public Shared Function SetItemAt(
                        listBoxName As Primitive,
                        index As Primitive,
                        value As Primitive
                   ) As Primitive

            Return ListBase.SetItemAt(listBoxName, index, value)
        End Function

        ''' <summary>
        ''' Adds an item to the end of the list.
        ''' </summary>
        ''' <param name="value">The item you want to add to the list. You can send an array to add all its items</param>
        ''' <returns>the index of the newly added item, or 0 if the operation failed.</returns>
        <ReturnValueType(VariableType.Double)>
        <ExMethod>
        Public Shared Function AddItem(listBoxName As Primitive, value As Primitive) As Primitive
            Return ListBase.AddItem(listBoxName, value)
        End Function

        ''' <summary>
        ''' Adds the given item to the list at the given index.
        ''' </summary>
        ''' <param name="value">the item you want to add to the list. You can send an array to add all its items</param>
        ''' <param name="index">The index you want to add the item at. The value of this index must be greater that 0 and less that list items count + 1, otherwise the item will not be added.</param>
        ''' <returns>True is the item is successfully added at the given index, otherwise False.</returns>
        <ExMethod>
        <ReturnValueType(VariableType.Boolean)>
        Public Shared Function AddItemAt(listBoxName As Primitive, value As Primitive, index As Primitive) As Primitive
            Return ListBase.AddItemAt(listBoxName, value, index)
        End Function

        ''' <summary>
        ''' Searches for the given value in the list, and removes the first found item.
        ''' </summary>
        ''' <param name="value">the item you want to remove. You can send an array to remove all its items</param>
        <ExMethod>
        <ReturnValueType(VariableType.Boolean)>
        Public Shared Function RemoveItem(listBoxName As Primitive, value As Primitive) As Primitive
            Return ListBase.RemoveItem(listBoxName, value)
        End Function

        ''' <summary>
        ''' Reomves all the items from the listbox
        ''' </summary>
        <ExMethod>
        Public Shared Sub RemoveAllItems(listBoxName As Primitive)
            ListBase.RemoveAllItems(listBoxName)
        End Sub

        ''' <summary>
        ''' Removes the list item that exists at the given index
        ''' </summary>
        ''' <param name="index">The index of the item you want to remove</param>
        ''' <returns>True if the item id removd, otherwise False.</returns>
        <ExMethod>
        <ReturnValueType(VariableType.Boolean)>
        Public Shared Function RemoveItemAt(listBoxName As Primitive, index As Primitive) As Primitive
            Return ListBase.RemoveItemAt(listBoxName, index)
        End Function


        ''' <summary>
        ''' Checkes whether or not the given item exists in the list.
        ''' </summary>
        ''' <param name="value">The value of the item to search for.</param>
        ''' <returns>True if the item found, or False otherwise.</returns>
        <ReturnValueType(VariableType.Boolean)>
        <ExMethod>
        Public Shared Function ContainsItem(listBoxName As Primitive, value As Primitive) As Primitive
            Return ListBase.ContainsItem(listBoxName, value)
        End Function

        ''' <summary>
        ''' Returns the index of the given item if it esists in the list, otherwise retruns 0.
        ''' </summary>
        ''' <param name="value">The item you want to find</param>
        ''' <returns>the index of the item if found, or 0 otherwise</returns>
        <ReturnValueType(VariableType.Double)>
        <ExMethod>
        Public Shared Function FindItem(listBoxName As Primitive, value As Primitive) As Primitive
            Return ListBase.FindItem(listBoxName, value)
        End Function

        ''' <summary>
        ''' Returns the index of the given item if it esists in the list in the given index range, otherwise retruns 0.
        ''' </summary>
        ''' <param name="value">The item you want to find.</param>
        ''' <param name="startIndex">The array index the search starts at.</param>
        ''' <param name="endIndex">The array index the search ends at. If endIndex is less than startIndex, the search direction will be reversed to find the last index of the item.</param>
        ''' <returns>the index of the item if found, or 0 otherwise.</returns>
        <ExMethod>
        <ReturnValueType(VariableType.Double)>
        Public Shared Function FindItemAt(listBoxName As Primitive, value As Primitive, startIndex As Primitive, endIndex As Primitive) As Primitive
            Return ListBase.FindItemAt(listBoxName, value, startIndex, endIndex)
        End Function

        ''' <summary>
        ''' This event is fired when the selected item in the list is changed.
        ''' </summary>
        Public Shared Custom Event OnSelection As SmallVisualBasicCallback
            AddHandler(handler As SmallVisualBasicCallback)
                Try
                    Dim name = [Event].SenderControl
                    Dim _sender = GetListBox(name)
                    Dim h = Sub(Sender As Wpf.Control, e As RoutedEventArgs)
                                [Event].HandleEvent(CType(Sender, FrameworkElement), e, handler)
                            End Sub

                    Control.RemovePrevEventHandler(
                            name,
                            NameOf(OnSelection),
                            Sub() RemoveHandler _sender.SelectionChanged, h
                     )
                    AddHandler _sender.SelectionChanged, h

                Catch ex As Exception
                    [Event].ShowErrorMessage(NameOf(OnSelection), ex)
                End Try

            End AddHandler

            RemoveHandler(handler As SmallVisualBasicCallback)
            End RemoveHandler

            RaiseEvent()
            End RaiseEvent
        End Event

    End Class
End Namespace