Option Explicit On


Imports Wpf = System.Windows.Controls
Imports Microsoft.SmallVisualBasic.Library
Imports App = Microsoft.SmallVisualBasic.Library.Internal.SmallBasicApplication

Namespace WinForms
    ''' <summary>
    ''' Represents the Menu control, which shows menu items on a bar, so the user can click any of them to drop down a list of sub menu items.
    ''' The form designer doesn't supoport adding a main menu at design time, but you can use the Form.AddMainMenu method to add it in runtime.
    ''' </summary>
    <SmallVisualBasicType>
    <HideFromIntellisense>
    Public NotInheritable Class MainMenu

        Private Shared Function GetMenu(menuName As String) As Wpf.Menu
            Dim c = Control.GetControl(menuName)
            Dim m = TryCast(c, Wpf.Menu)
            If m Is Nothing Then
                Throw New Exception($"{menuName} is not a name of a Menu.")
            End If
            Return m
        End Function

        ''' <summary>
        ''' Adds a menu item to the current menu.
        ''' </summary>
        ''' <param name="itemName">The name of the new menu item</param>
        ''' <param name="text">The title of the menu item</param>
        ''' <param name="shortcut">The keyboard shortcut keys, like Ctrl+N, or use "" if there is no shortcut.</param>
        ''' <returns>The menu item that have been added.</returns>
        <ExMethod>
        <ReturnValueType(VariableType.MenuItem)>
        Public Shared Function AddItem(
                        menuName As Primitive,
                        itemName As Primitive,
                        text As Primitive,
                        shortcut As Primitive
                    ) As Primitive

            Dim _menuName = menuName.AsString()
            Dim key = Form.ValidateArgs(_menuName.Substring(0, _menuName.IndexOf(".")), itemName.AsString())

            App.Invoke(
                Sub()
                    Try
                        Dim memu = GetMenu(menuName)
                        Dim item = New Wpf.MenuItem() With {
                            .InputGestureText = shortcut,
                            .Header = text,
                            .Name = itemName
                        }
                        memu.Items.Add(item)
                        Forms._controls(key) = item

                    Catch ex As Exception
                        Control.ReportSubError(menuName, "AddItem", ex)
                    End Try
                End Sub)

            Return New Primitive(key)
        End Function

        ''' <summary>
        ''' Removes a menu item from the current menu.
        ''' </summary>
        ''' <param name="itemName">The name of the new menu item</param>
        <ExMethod>
        Public Shared Sub RemoveItem(
                        menuName As Primitive,
                        itemName As Primitive
                    )

            App.Invoke(
                Sub()
                    Try
                        Dim menu = GetMenu(menuName)
                        RemoveMenuItem(menu.Items, itemName.ToString().ToLower())
                    Catch ex As Exception
                        Control.ReportSubError(menuName, "RemoveItem", ex)
                    End Try
                End Sub)
        End Sub

        Friend Shared Sub RemoveMenuItem(items As Wpf.ItemCollection, itemName As String)
            If Not Forms._controls.ContainsKey(itemName) Then Return

            Dim item = CType(Forms._controls(itemName), Wpf.MenuItem)
            Dim subItems = item.Items
            Dim formName = itemName.Substring(0, itemName.IndexOf(".") + 1)

            For i = subItems.Count - 1 To 0 Step -1
                Dim m = CType(subItems(i), Wpf.MenuItem)
                RemoveMenuItem(subItems, formName + m.Name.ToLower())
            Next

            items.Remove(item)
            Forms._controls.Remove(itemName)
        End Sub

        ''' <summary>
        ''' Returns an array containing the child menu items of the current menu.
        ''' </summary>
        ''' <returns>an array of menu items</returns>
        <ReturnValueType(VariableType.Array)>
        <ExProperty>
        Public Shared Function GetMenuItems(menuName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        Dim name = menuName.AsString()
                        Dim formName = name.Substring(0, name.IndexOf("."))
                        Dim items As New Dictionary(Of Primitive, Primitive)
                        Dim i = 1

                        For Each item In GetMenu(menuName).Items
                            Dim m = TryCast(item, Wpf.MenuItem)
                            If m IsNot Nothing Then
                                items.Add(i, New Primitive($"{formName}.{m.Name}"))
                                i += 1
                            End If
                        Next

                        GetMenuItems = New Primitive() With {.ArrayMap = items}

                    Catch ex As Exception
                        Control.ReportError(menuName, "MenuItems", ex)
                    End Try
                End Sub)
        End Function

    End Class

End Namespace
