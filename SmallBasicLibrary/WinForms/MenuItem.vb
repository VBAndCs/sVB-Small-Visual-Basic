Option Explicit On

Imports Wpf = System.Windows.Controls
Imports Microsoft.SmallVisualBasic.Library
Imports App = Microsoft.SmallVisualBasic.Library.Internal.SmallBasicApplication
Imports System.Windows

Namespace WinForms
    ''' <summary>
    ''' Represents the MenuItem control, which shows a menuItem on the main menu bar or on the dropdown list of a parent manuItem.
    ''' The user can click the menu item to perform the task you programmed in the OnClick event handler.
    ''' You can also set the Checkable property to True to allow the user to check or uncheck the menu item, hence you can use the Checked property and the OnCheck event to respond the user choices.
    ''' The form designer doesn't supoport adding menu items at design time, but you can use the MainMenu.AddItem method to add an item to the main menu, or use the MenuItem.AddItem to add an item to a pearent menu item.
    ''' You can also use the MenuItem.AddSeparator to add a separator line to the menu item.
    ''' </summary>
    <SmallVisualBasicType>
    <HideFromIntellisense>
    Public NotInheritable Class MenuItem

        Friend Shared MenuEventHandlers As New Dictionary(Of String, RoutedEventHandler)

        Friend Shared Function GetMenuItem(itemName As String) As Wpf.MenuItem
            Dim c = Control.GetControl(itemName)
            Dim m = TryCast(c, Wpf.MenuItem)
            If m Is Nothing Then
                Throw New Exception($"{itemName} is not a name of a MenuItem.")
            End If
            Return m
        End Function

        ''' <summary>
        ''' Adds a sub menu item to the current menu item
        ''' </summary>
        ''' <param name="itemName">The name of the new sub item</param>
        ''' <param name="text">The title of the sub menu item</param>
        ''' <param name="shortcut">The keyboard shortcut keys, like Ctrl+N</param>
        ''' <returns>The menu item that have been added.</returns>
        <ExMethod>
        <ReturnValueType(VariableType.MenuItem)>
        Public Shared Function AddItem(
                        parentItem As Primitive,
                        itemName As Primitive,
                        text As Primitive,
                        shortcut As Primitive
                    ) As Primitive

            Dim _parentName = parentItem.AsString
            Dim key = Form.ValidateArgs(_parentName.Substring(0, _parentName.IndexOf(".")), itemName.AsString())

            App.Invoke(
                Sub()
                    Try
                        Dim memuItem = GetMenuItem(parentItem)
                        Dim item As New Wpf.MenuItem() With {
                            .InputGestureText = shortcut,
                            .Header = text,
                            .Name = itemName
                        }
                        memuItem.Items.Add(item)
                        Forms._controls(key) = item

                    Catch ex As Exception
                        Control.ReportSubError(parentItem, "AddItem", ex)
                    End Try
                End Sub)

            Return New Primitive(key)
        End Function

        ''' <summary>
        ''' Removes a sub menu item from the current menu item.
        ''' </summary>
        ''' <param name="itemName">The name of the new menu item</param>
        <ExMethod>
        Public Shared Sub RemoveItem(
                        parentItem As Primitive,
                        itemName As Primitive
                    )

            App.Invoke(
                Sub()
                    Try
                        Dim menuItem = GetMenuItem(parentItem)
                        MainMenu.RemoveMenuItem(menuItem.Items, itemName.ToString().ToLower())
                    Catch ex As Exception
                        Control.ReportSubError(parentItem, "RemoveItem", ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Adds a separator line betwen menu items
        ''' </summary>
        <ExMethod>
        Public Shared Sub AddSeparator(menuItemName As Primitive)
            Dim _parentName = menuItemName.AsString

            App.Invoke(
                Sub()
                    Try
                        Dim memuItem = GetMenuItem(menuItemName)
                        Dim item As New Wpf.Separator()
                        memuItem.Items.Add(item)

                    Catch ex As Exception
                        Control.ReportSubError(menuItemName, "AddSeparator", ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' When it is True, user can interact with the control.
        ''' When it is False,  the control is disabled, and user can't interact with it.
        ''' </summary>
        <ReturnValueType(VariableType.Boolean)>
        <ExProperty>
        Public Shared Function GetCheckable(menuItemName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetCheckable = GetMenuItem(menuItemName).IsCheckable
                    Catch ex As Exception
                        Control.ReportError(menuItemName, "Checkable", ex)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetCheckable(menuItemName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        GetMenuItem(menuItemName).IsCheckable = value
                    Catch ex As Exception
                        Control.ReportPropertyError(menuItemName, "Checkable", value, ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' When True, the menu item is checked.
        ''' When False, the menu item is unchecked.
        ''' Set the Checkable property to True, to allow the user to check the menu item.
        ''' </summary>
        <ReturnValueType(VariableType.Boolean)>
        <ExProperty>
        Public Shared Function GetChecked(menuItemName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetChecked = GetMenuItem(menuItemName).IsChecked
                    Catch ex As Exception
                        Control.ReportError(menuItemName, "Checked", ex)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetChecked(menuItemName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        GetMenuItem(menuItemName).IsChecked = value
                    Catch ex As Exception
                        Control.ReportPropertyError(menuItemName, "Checked", value, ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Returns as array containing the child menu items of the current menu item.
        ''' Separators will not be included.
        ''' </summary>
        ''' <returns>an array of menu items</returns>
        <ReturnValueType(VariableType.Array)>
        <ExProperty>
        Public Shared Function GetMenuItems(menuItemName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        Dim name = menuItemName.AsString()
                        Dim formName = name.Substring(0, name.IndexOf("."))
                        Dim items As New Dictionary(Of Primitive, Primitive)
                        Dim i = 1

                        For Each item In GetMenuItem(menuItemName).Items
                            Dim m = TryCast(item, Wpf.MenuItem)
                            If m IsNot Nothing Then
                                items.Add(i, New Primitive($"{formName}.{m.Name}"))
                                i += 1
                            End If
                        Next

                        GetMenuItems = New Primitive() With {.ArrayMap = items}

                    Catch ex As Exception
                        Control.ReportError(menuItemName, "MenuItems", ex)
                    End Try
                End Sub)
        End Function

        Friend Shared ShortcutHandlers As New Dictionary(Of Wpf.MenuItem, ShortcutHandler)

        ''' <summary>
        ''' Fired when user clicks the menu item by the left mouse button.
        ''' </summary>
        Public Shared Custom Event OnClick As SmallVisualBasicCallback
            AddHandler(handler As SmallVisualBasicCallback)
                App.Invoke(
                    Sub()
                        Dim _sender = GetMenuItem([Event].SenderControl)
                        Dim shortcut = _sender.InputGestureText
                        If shortcut <> "" Then
                            ShortcutHandlers(_sender) = New ShortcutHandler(shortcut, handler)
                        End If

                        Dim h = Sub(Sender As Object, e As System.Windows.RoutedEventArgs)
                                    [Event].HandleEvent(CType(Sender, System.Windows.FrameworkElement), e, handler)
                                End Sub

                        Control.RemovePrevEventHandler(
                            [Event].SenderControl,
                            NameOf(OnClick),
                            Sub()
                                RemoveHandler _sender.Click, h
                                ShortcutHandlers(_sender) = Nothing
                            End Sub
                        )
                        AddHandler _sender.Click, h
                        Dim key = (Control.senderAssembly & ":" & [Event].SenderControl.AsString() & ".OnClick").ToLower()
                        MenuEventHandlers(key) = h
                    End Sub)
            End AddHandler

            RemoveHandler(handler As SmallVisualBasicCallback)
            End RemoveHandler

            RaiseEvent()
            End RaiseEvent
        End Event


        ''' <summary>
        ''' Fired when the submenu is opened.
        ''' </summary>
        Public Shared Custom Event OnOpen As SmallVisualBasicCallback
            AddHandler(handler As SmallVisualBasicCallback)
                Try
                    Dim _sender = GetMenuItem([Event].SenderControl)
                    Dim h = Sub(Sender As Wpf.Control, e As System.Windows.RoutedEventArgs)
                                [Event].HandleEvent(CType(Sender, System.Windows.FrameworkElement), e, handler)
                            End Sub

                    Control.RemovePrevEventHandler(
                        [Event].SenderControl,
                        NameOf(OnOpen),
                        Sub() RemoveHandler _sender.SubmenuOpened, h
                    )
                    AddHandler _sender.SubmenuOpened, h
                    Dim key = (Control.senderAssembly & ":" & [Event].SenderControl.AsString() & ".OnOpen").ToLower()
                    MenuEventHandlers(key) = h

                Catch ex As Exception
                    [Event].ShowErrorMessage(NameOf(OnOpen), ex)
                End Try
            End AddHandler

            RemoveHandler(handler As SmallVisualBasicCallback)
            End RemoveHandler

            RaiseEvent()
            End RaiseEvent
        End Event

        ''' <summary>
        ''' Fired when the checked state is changed.
        ''' </summary>
        Public Shared Custom Event OnCheck As SmallVisualBasicCallback
            AddHandler(handler As SmallVisualBasicCallback)
                App.Invoke(
                    Sub()
                        Try
                            Dim _sender = GetMenuItem([Event].SenderControl)
                            Dim shortcut = _sender.InputGestureText
                            If shortcut <> "" AndAlso _sender.IsCheckable Then
                                ShortcutHandlers(_sender) = New ShortcutHandler(shortcut, Nothing)
                            End If

                            Dim h = Sub(Sender As Wpf.Control, e As System.Windows.RoutedEventArgs)
                                        [Event].HandleEvent(Sender, e, handler)
                                    End Sub

                            Control.RemovePrevEventHandler(
                                    [Event].SenderControl,
                                    NameOf(OnCheck),
                                    Sub()
                                        RemoveHandler _sender.Checked, h
                                        RemoveHandler _sender.Unchecked, h
                                    End Sub
                            )
                            AddHandler _sender.Checked, h
                            AddHandler _sender.Unchecked, h
                            Dim key = (Control.senderAssembly & ":" & [Event].SenderControl.AsString() & ".OnCheck").ToLower()
                            MenuEventHandlers(key) = h

                        Catch ex As Exception
                            [Event].ShowErrorMessage(NameOf(OnCheck), ex)
                        End Try
                    End Sub)
            End AddHandler

            RemoveHandler(handler As SmallVisualBasicCallback)
            End RemoveHandler

            RaiseEvent()
            End RaiseEvent
        End Event

    End Class

End Namespace
