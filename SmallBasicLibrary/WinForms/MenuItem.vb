Option Explicit On

Imports Wpf = System.Windows.Controls
Imports Microsoft.SmallVisualBasic.Library
Imports App = Microsoft.SmallVisualBasic.Library.Internal.SmallBasicApplication

Namespace WinForms
    <SmallBasicType>
    <HideFromIntellisense>
    Public NotInheritable Class MenuItem
        Private Shared Function GetMenuItem(itemName As String) As Wpf.MenuItem
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
            Dim key = Form.ValidateArgs(_parentName.Substring(0, _parentName.IndexOf(".")), itemName)

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

            Return key
        End Function

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
        ''' Set the Chickable property to True, to allow the user to check the menu item.
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
                                items.Add(i, $"{formName}.{m.Name}")
                                i += 1
                            End If
                        Next

                        GetMenuItems = New Primitive() With {._arrayMap = items}

                    Catch ex As Exception
                        Control.ReportError(menuItemName, "MenuItems", ex)
                    End Try
                End Sub)
        End Function

        ''' <summary>
        ''' Fired when the submenu is opened.
        ''' </summary>
        Public Shared Custom Event OnOpen As SmallBasicCallback
            AddHandler(handler As SmallBasicCallback)
                Try
                    Dim _sender = GetMenuItem([Event].SenderControl)
                    AddHandler _sender.SubmenuOpened, Sub(Sender As Wpf.Control, e As System.Windows.RoutedEventArgs) [Event].EventsHandler(CType(Sender, System.Windows.FrameworkElement), e, handler)
                Catch ex As Exception
                    [Event].ShowErrorMessage(NameOf(OnOpen), ex)
                End Try
            End AddHandler

            RemoveHandler(handler As SmallBasicCallback)
            End RemoveHandler

            RaiseEvent()
            End RaiseEvent
        End Event

        ''' <summary>
        ''' Fired when the checked state is changed.
        ''' </summary>
        Public Shared Custom Event OnCheck As SmallBasicCallback
            AddHandler(handler As SmallBasicCallback)
                Try
                    Dim _sender = GetMenuItem([Event].SenderControl)
                    Dim _handler = Sub(Sender As Wpf.Control, e As System.Windows.RoutedEventArgs) [Event].EventsHandler(Sender, e, handler)
                    AddHandler _sender.Checked, _handler
                    AddHandler _sender.Unchecked, _handler
                Catch ex As Exception
                    [Event].ShowErrorMessage(NameOf(OnCheck), ex)
                End Try
            End AddHandler

            RemoveHandler(handler As SmallBasicCallback)
            End RemoveHandler

            RaiseEvent()
            End RaiseEvent
        End Event

    End Class

End Namespace
