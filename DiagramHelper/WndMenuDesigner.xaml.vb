Public Class WndMenuDesigner

    Dim newID As Integer = 0

    Private Sub WndMenuDesigner_ContentRendered(sender As Object, e As EventArgs) Handles Me.ContentRendered
        _CurrrentMenuItem = MainMenu
        MainMenu.Name = "MainMenu1"
        newID = GetNewID(MainMenu)

        If MainMenu.Items.Count = 0 Then
            BtnAddNext_Click(Nothing, Nothing)
        Else
            AddDesignProperties(MainMenu)
            CurrrentMenuItem = MainMenu.Items(0)
            Dim m = TryCast(_CurrrentMenuItem, MenuItem)
            If m IsNot Nothing Then m.IsSubmenuOpen = True
        End If

        CmbKeys.ItemsSource = {
            "(None)",
            "-", "+", "[", "]",
            "0", "1", "2", "3", "4", "5", "6", "7", "8", "9",
            "A", "B", "Back", "C", "D", "Delete", "Down", "E", "Esc",
            "F", "F1", "F2", "F3", "F4", "F5", "F6", "F7", "F8", "F9", "F10", "F11", "F12",
            "G", "H", "I", "J", "K", "L", "Left", "M", "N", "O", "P", "Q", "R", "Right", "S", "T", "Tab", "U", "Up", "V", "W", "X", "Y", "Z"
        }
        CmbKeys.SelectedIndex = 0
    End Sub

    Private Sub AddDesignProperties(parent As ItemsControl)
        For Each item In parent.Items
            Dim m = TryCast(item, MenuItem)
            If m Is Nothing Then Continue For
            m.StaysOpenOnClick = True
            AddHandler m.PreviewMouseDown, AddressOf MenuItem_Click
            AddDesignProperties(m)
        Next
    End Sub

    Dim _CurrrentMenuItem As Control

    Private Property CurrrentMenuItem As Control
        Get
            Return _CurrrentMenuItem
        End Get

        Set(value As Control)
            If Not CommintChanges() Then Return

            If _CurrrentMenuItem IsNot Nothing Then
                _CurrrentMenuItem.ClearValue(BackgroundProperty)
                _CurrrentMenuItem.ClearValue(BorderBrushProperty)
                _CurrrentMenuItem.ClearValue(BorderThicknessProperty)
            End If

            _CurrrentMenuItem = value
            If value IsNot Nothing Then SetPropFields()

        End Set
    End Property

    Private Sub SetPropFields()
        _CurrrentMenuItem.Background = Brushes.Yellow
        _CurrrentMenuItem.BorderBrush = Brushes.Black
        _CurrrentMenuItem.BorderThickness = New Thickness(1)
        TxtName.Text = _CurrrentMenuItem.Name

        Dim m = TryCast(_CurrrentMenuItem, MenuItem)
        If m Is Nothing Then
            TxtText.IsEnabled = False
            TxtText.Text = ""

            GrpShortcut.IsEnabled = False
            ChkCtrl.IsChecked = False
            ChkShift.IsChecked = False
            ChkAlt.IsChecked = False
            CmbKeys.SelectedIndex = 0

            GrpCheck.IsEnabled = False
            ChkCheckable.IsChecked = False
            ChkChecked.IsChecked = False

        Else
            TxtText.IsEnabled = True
            TxtText.Text = m.Header?.ToString()
            If m.Items.Count = 0 Then
                GrpShortcut.IsEnabled = True
                GrpCheck.IsEnabled = True
                ChkCheckable.IsChecked = (LCase(m.Tag) = "true")
                ChkChecked.IsChecked = m.IsChecked
            Else
                GrpShortcut.IsEnabled = False
                GrpCheck.IsEnabled = False
                ChkCheckable.IsChecked = False
                ChkChecked.IsChecked = False
            End If

            Dim s = m.InputGestureText
            If s = "" Then
                ChkCtrl.IsChecked = False
                ChkShift.IsChecked = False
                ChkAlt.IsChecked = False
                CmbKeys.SelectedIndex = 0
            Else
                ChkCtrl.IsChecked = s.Contains("Ctrl+")
                ChkShift.IsChecked = s.Contains("Shift+")
                ChkAlt.IsChecked = s.Contains("Alt+")
                CmbKeys.SelectedItem = s.Substring(s.LastIndexOf("+") + 1)
            End If
        End If

        SelectName()
    End Sub

    Private Function CommintChanges() As Boolean
        Dim activeControl = FocusManager.GetFocusedElement(Me)
        If activeControl Is TxtName Then
            If Not CommitName() Then Return False
        ElseIf activeControl Is TxtText Then
            If Not CommitText() Then Return False
        End If
        Return True
    End Function

    Private Sub BtnAddNext_Click(sender As Object, e As RoutedEventArgs)
        Dim newItem = CreateMenuItem()
        Dim index = 0
        Dim parent = _CurrrentMenuItem

        If parent IsNot MainMenu Then
            Dim p = CType(parent.Parent, ItemsControl)
            index = p.Items.IndexOf(parent) + 1
            parent = p
        End If

        Dim m = CType(parent, ItemsControl)
        m.Items.Insert(index, newItem)
        OpenMenu(m)
        ' Set the property not the back filed to change colors
        CurrrentMenuItem = newItem
        TxtName.Focus()
    End Sub

    Private Function CreateMenuItem() As MenuItem
        Dim m = New MenuItem() With {
            .Name = MnuPrefix & newID,
            .Header = "MenuItem" & newID,
            .StaysOpenOnClick = True
        }
        AddHandler m.PreviewMouseDown, AddressOf MenuItem_Click
        newID += 1
        Return m
    End Function

    Private Const MnuPrefix = "Mnu"
    Private Const SepPrefix = "Sep"

    Private Function GetNewID(parent As ItemsControl, Optional id As Integer = 1) As Integer
        Dim mp = MnuPrefix.ToLower()
        Dim sp = SepPrefix.ToLower()
        Dim L = mp.Length

        For Each m As Control In parent.Items
            Dim name = m.Name.ToLower()
            If name.StartsWith(mp) OrElse name.StartsWith(sp) Then
                Dim n = name.Substring(L)
                If IsNumeric(n) Then
                    If CInt(n) >= id Then id = CInt(n) + 1
                End If
            End If

            If TypeOf m IsNot Separator Then id = GetNewID(m, id)
        Next
        Return id
    End Function


    Private Sub OpenMenu(menu As Control)
        If menu Is MainMenu Then Return

        Dim m = TryCast(menu, MenuItem)
        If m Is Nothing Then
            m = TryCast(menu.Parent, MenuItem)
            If m Is Nothing Then Return
        End If

        Dim menus As New List(Of MenuItem)
        Do
            menus.Add(m)
            Dim p = m.Parent
            If p Is MainMenu Then Exit Do
            m = p
        Loop

        For i = menus.Count - 1 To 0 Step -1
            menus(i).IsSubmenuOpen = True
        Next
    End Sub

    Private Sub BtnAddChild_Click(sender As Object, e As RoutedEventArgs)
        If TypeOf _CurrrentMenuItem Is Separator Then
            Beep()
            OpenMenu(_CurrrentMenuItem)
        Else
            Dim newItem = CreateMenuItem()
            Dim m = CType(_CurrrentMenuItem, ItemsControl)
            m.Items.Add(newItem)
            ClearShortcutandCkecked(m)
            OpenMenu(_CurrrentMenuItem)
            ' Set the property not the back filed to change colors
            CurrrentMenuItem = newItem
            TxtName.Focus()
        End If
    End Sub

    Private Shared Sub ClearShortcutAndCkecked(item As ItemsControl)
        Dim m = TryCast(item, MenuItem)
        If m IsNot Nothing Then
            m.InputGestureText = ""
            m.IsChecked = False
            m.Tag = ""
        End If
    End Sub

    Private Sub MenuItem_Click(sender As Object, e As RoutedEventArgs)
        ' Set the property not the back filed to change colors
        CurrrentMenuItem = sender
    End Sub

    Private Sub Window_PreviewKeyDown(sender As Object, e As KeyEventArgs)
        Dim activeControl = FocusManager.GetFocusedElement(Me)
        Dim cancelLeftAndRight = TypeOf activeControl Is TextBox
        Dim cancelUpAndDown = TypeOf activeControl Is ComboBox OrElse
                                                    TypeOf activeControl Is ComboBoxItem

        Select Case e.Key
            Case Key.Escape
                DontAskAgain = True

            Case Key.F3
                BtnAddNext_Click(Nothing, Nothing)
                e.Handled = True

            Case Key.F4
                BtnAddChild_Click(Nothing, Nothing)
                e.Handled = True

            Case Key.Up
                If e.KeyboardDevice.Modifiers = ModifierKeys.Control Then
                    DecreaseIndex()
                ElseIf cancelUpAndDown Then
                    Return
                Else
                    SelectPrevItem()
                End If
                e.Handled = True

            Case Key.Down
                If e.KeyboardDevice.Modifiers = ModifierKeys.Control Then
                    IncreaseIndex()
                ElseIf cancelUpAndDown Then
                    Return
                ElseIf _CurrrentMenuItem.Parent Is MainMenu Then
                    SelectChild()
                Else
                    SelectNextItem()
                End If
                e.Handled = True

            Case Key.Left
                If e.KeyboardDevice.Modifiers = ModifierKeys.Control Then
                    DecreaseDepth()
                ElseIf cancelLeftAndRight Then
                    Return
                ElseIf _CurrrentMenuItem.Parent Is MainMenu Then
                    SelectPrevItem()
                Else
                    SelectParent()
                End If
                e.Handled = True

            Case Key.Right
                If e.KeyboardDevice.Modifiers = ModifierKeys.Control Then
                    IncreaseDepth()
                ElseIf cancelLeftAndRight Then
                    Return
                ElseIf _CurrrentMenuItem.Parent Is MainMenu Then
                    SelectNextItem()
                Else
                    SelectChild()
                End If
                e.Handled = True
        End Select
    End Sub

    Private Sub IncreaseDepth()
        If _CurrrentMenuItem Is MainMenu Then
            Beep()
            Return
        End If

        Dim parent = CType(_CurrrentMenuItem.Parent, ItemsControl)
        Dim items = parent.Items
        Dim i = items.IndexOf(_CurrrentMenuItem)
        If i = 0 Then
            Beep()
            Return
        End If

        Dim newParenr = items(i - 1)
        items.Remove(_CurrrentMenuItem)
        newParenr.items.add(_CurrrentMenuItem)
        ClearShortcutandCkecked(newParenr)
        Helper.UpdateControl(newParenr)
        OpenMenu(_CurrrentMenuItem)
    End Sub

    Private Sub DecreaseDepth()
        If _CurrrentMenuItem Is MainMenu Then
            Beep()
            Return
        End If

        Dim parent = CType(_CurrrentMenuItem.Parent, ItemsControl)
        If parent Is MainMenu Then
            Beep()
            Return
        End If

        Dim grandParent = CType(parent.Parent, ItemsControl)
        Dim items = grandParent.Items
        Dim i = items.IndexOf(parent)
        parent.Items.Remove(_CurrrentMenuItem)
        items.Insert(i + 1, _CurrrentMenuItem)
        CType(parent, MenuItem).IsSubmenuOpen = False
        OpenMenu(_CurrrentMenuItem)
    End Sub

    Private Sub IncreaseIndex()
        If _CurrrentMenuItem Is MainMenu Then
            Beep()
            Return
        End If

        Dim parent = CType(_CurrrentMenuItem.Parent, ItemsControl)
        Dim items = parent.Items
        Dim i = items.IndexOf(_CurrrentMenuItem)
        If i = items.Count - 1 Then
            Beep()
            Return
        End If

        items.RemoveAt(i)
        items.Insert(i + 1, _CurrrentMenuItem)
        OpenMenu(_CurrrentMenuItem)
    End Sub

    Private Sub DecreaseIndex()
        If _CurrrentMenuItem Is MainMenu Then
            Beep()
            Return
        End If

        Dim parent = CType(_CurrrentMenuItem.Parent, ItemsControl)
        Dim items = parent.Items
        Dim i = items.IndexOf(_CurrrentMenuItem)
        If i = 0 Then
            Beep()
            Return
        End If

        items.RemoveAt(i)
        items.Insert(i - 1, _CurrrentMenuItem)
        OpenMenu(_CurrrentMenuItem)
    End Sub

    Private Sub SelectPrevItem()
        If _CurrrentMenuItem Is MainMenu Then
            Beep()
            Return
        End If

        Dim parent = CType(_CurrrentMenuItem.Parent, ItemsControl)
        Dim items = parent.Items
        Dim i = items.IndexOf(_CurrrentMenuItem)
        If i > 0 Then
            CurrrentMenuItem = items(i - 1)
            OpenMenu(_CurrrentMenuItem)
        ElseIf parent.Parent Is MainMenu Then
            SelectParent()
        Else
            Beep()
        End If
    End Sub

    Private Sub SelectNextItem()
        If _CurrrentMenuItem Is MainMenu Then
            Beep()
            Return
        End If

        Dim items = CType(_CurrrentMenuItem.Parent, ItemsControl).Items
        Dim i = items.IndexOf(_CurrrentMenuItem)
        If i = items.Count - 1 Then
            Beep()
            Return
        End If

        CurrrentMenuItem = items(i + 1)
        OpenMenu(_CurrrentMenuItem)
    End Sub

    Private Sub SelectChild()
        Dim m = TryCast(_CurrrentMenuItem, ItemsControl)
        If m Is Nothing Then
            Beep()
            Return
        End If

        If m.Items.Count = 0 Then
            Beep()
            Return
        End If

        OpenMenu(m)
        CurrrentMenuItem = m.Items(0)
        OpenMenu(m)
    End Sub

    Private Sub SelectParent()
        If _CurrrentMenuItem Is MainMenu Then
            Beep()
            Return
        End If

        Dim parent = _CurrrentMenuItem.Parent
        If parent Is MainMenu Then
            Beep()
            Return
        End If
        CurrrentMenuItem = parent
        OpenMenu(_CurrrentMenuItem)
    End Sub


    Private Sub TxtName_PreviewKeyDown(sender As Object, e As KeyEventArgs)
        Select Case e.Key
            Case Key.Enter
                CommitName()
                e.Handled = True

            Case Key.Escape
                TxtName.Text = CType(_CurrrentMenuItem, Control).Name
                e.Handled = True
        End Select
    End Sub

    Private Function CommitName() As Boolean
        Dim newName = TxtName.Text.Trim()
        If newName = "" Then Return False
        If CurrrentMenuItem.Name = newName Then Return True

        If newName = "" Then
            Beep()
            Return False
        End If

        newName = newName(0).ToString().ToUpper & If(newName.Length > 1, newName.Substring(1), "")
        If TxtName.Text <> newName Then TxtName.Text = newName
        If Designer.CurrentPage.SetControlName(_CurrrentMenuItem, newName, MainMenu) Then
            If TxtText.IsEnabled AndAlso newName.ToLower().EndsWith(TxtText.Text.ToLower()) Then
                CommitText()
            End If
            Return True
        End If

        Return False
    End Function

    Private Sub BtnDelete_Click(sender As Object, e As RoutedEventArgs)
        If _CurrrentMenuItem Is MainMenu Then
            Beep()
            Return
        End If

        Dim parent = CType(_CurrrentMenuItem.Parent, ItemsControl)
        Dim items = parent.Items
        Dim i = items.IndexOf(_CurrrentMenuItem)
        items.RemoveAt(i)
        If i = items.Count Then
            If i = 0 Then
                CurrrentMenuItem = parent
            Else
                CurrrentMenuItem = items(i - 1)
            End If
        Else
            CurrrentMenuItem = items(i)
        End If
        OpenMenu(_CurrrentMenuItem)
    End Sub

    Private Sub TxtName_LostFocus(sender As Object, e As RoutedEventArgs)
        If Not CommitName() Then
            e.Handled = True
            Dispatcher.BeginInvoke(
                   Windows.Threading.DispatcherPriority.Background,
                   Sub() TxtName.Focus()
            )
        End If
    End Sub

    Private Sub BtnAddSep_Click(sender As Object, e As RoutedEventArgs)
        Dim newItem As New Separator()
        newItem.Name = SepPrefix & newID
        newID += 1
        AddHandler newItem.PreviewMouseDown, AddressOf MenuItem_Click

        Dim index = 0
        Dim parent = _CurrrentMenuItem

        If parent IsNot MainMenu Then
            Dim p = CType(parent.Parent, ItemsControl)
            index = p.Items.IndexOf(parent) + 1
            parent = p
        End If

        CType(parent, ItemsControl).Items.Insert(index, newItem)
        CurrrentMenuItem = newItem
        OpenMenu(parent)
    End Sub

    Private Sub TxtText_PreviewKeyDown(sender As Object, e As KeyEventArgs)
        Select Case e.Key
            Case Key.Enter
                CommitText()
                e.Handled = True

            Case Key.Escape
                Dim m = TryCast(_CurrrentMenuItem, MenuItem)
                If m Is Nothing Then
                    Beep()
                Else
                    TxtText.Text = m.Header
                End If
                e.Handled = True
        End Select
    End Sub

    Private Function CommitText() As Boolean
        Dim m = TryCast(_CurrrentMenuItem, MenuItem)
        If m Is Nothing Then
            Beep()
        Else
            Dim text = Trim(TxtText.Text)
            TxtText.Text = text

            If text = "" Then
                MsgBox("The menu item text must contain at least one character.")
            Else
                m.Header = text
                Return True
            End If
        End If

        Return False
    End Function

    Private Sub TxtText_LostFocus(sender As Object, e As RoutedEventArgs)
        If Not CommitText() Then
            e.Handled = True
            Dispatcher.BeginInvoke(
                   Windows.Threading.DispatcherPriority.Background,
                   Sub() TxtText.Focus()
            )
        End If
    End Sub

    Private Sub Controls_GotFocus(sender As Object, e As RoutedEventArgs)
        Helper.RunLater(Me, Sub() OpenMenu(_CurrrentMenuItem), 200)
    End Sub

    Private Sub ChkCheckable_Checked(sender As Object, e As RoutedEventArgs)
        _CurrrentMenuItem.Tag = ChkCheckable.IsChecked
    End Sub

    Private Sub ChkChecked_Checked(sender As Object, e As RoutedEventArgs)
        If GrpCheck.IsEnabled Then
            CType(_CurrrentMenuItem, MenuItem).IsChecked = ChkChecked.IsChecked
        End If
    End Sub

    Private Sub OnShortcutChanged(sender As Object, e As RoutedEventArgs)
        If Not GrpShortcut.IsEnabled Then Return

        If CmbKeys.SelectedIndex = 0 Then
            CType(_CurrrentMenuItem, MenuItem).InputGestureText = ""
        Else
            Dim s = ""
            If ChkCtrl.IsChecked Then s = "Ctrl+"
            If ChkShift.IsChecked Then s &= "Shift+"
            If ChkAlt.IsChecked Then s &= "Alt+"
            CType(_CurrrentMenuItem, MenuItem).InputGestureText = s & CmbKeys.SelectedItem
        End If
        Controls_GotFocus(sender, Nothing)
    End Sub


    Dim DontAskAgain As Boolean = False
    Dim ForceClose As Boolean = False

    Private Sub BtnCancel_Click(sender As Object, e As System.Windows.RoutedEventArgs)
        If MsgBox("You are about to close the window without saving the changes.",
                  MsgBoxStyle.OkCancel Or MsgBoxStyle.DefaultButton2 Or MsgBoxStyle.Exclamation,
                  "Be Careful!") = MsgBoxResult.Ok Then
            ForceClose = True
            Me.DialogResult = False
        End If
    End Sub

    Private Sub BtnOk_Click(sender As Object, e As System.Windows.RoutedEventArgs)
        If Not CommintChanges() Then Return

        CurrrentMenuItem = Nothing
        RemoveDesignPropertis(MainMenu)
        ForceClose = True
        Me.DialogResult = True
    End Sub

    Private Sub RemoveDesignPropertis(menu As ItemsControl)
        Dim m = TryCast(menu, MenuItem)
        If m IsNot Nothing Then m.StaysOpenOnClick = False

        For Each item In menu.Items
            If TypeOf item IsNot Separator Then RemoveDesignPropertis(item)
        Next
    End Sub

    Private Sub Window_Closing(sender As Object, e As ComponentModel.CancelEventArgs)
        If ForceClose Then Return

        If DontAskAgain Then
            e.Cancel = True
            DontAskAgain = False
            Return
        End If

        e.Cancel = MsgBox(
                             "You are about to close the window without saving the changes.",
                             MsgBoxStyle.OkCancel Or MsgBoxStyle.DefaultButton2 Or MsgBoxStyle.Exclamation,
                             "Be Careful!"
                         ) = MsgBoxResult.Cancel
    End Sub

    Private Sub TxtText_GotFocus(sender As Object, e As RoutedEventArgs)
        Helper.RunLater(Me, Sub() CType(sender, TextBox).SelectAll(), 1)
        Controls_GotFocus(sender, Nothing)
    End Sub

    Private Sub TxtName_GotFocus(sender As Object, e As RoutedEventArgs)
        Helper.RunLater(Me, AddressOf SelectName, 1)
    End Sub

    Dim isDefaultText As Boolean

    Private Sub SelectName()
        Dim name = TxtName.Text.ToLower()
        Dim len = TxtName.Text.Length
        If name.StartsWith("menu") Then
            TxtName.Select(4, len - 4)
        ElseIf name.StartsWith("mnu") Then
            TxtName.Select(3, len - 3)
        Else
            TxtName.SelectAll()
        End If

        Controls_GotFocus(TxtName, Nothing)
        isDefaultText = TxtText.Text.ToLower.StartsWith("menuitem")
    End Sub

    Private Sub TxtName_TextChanged(sender As Object, e As TextChangedEventArgs)
        If Not TxtText.IsEnabled OrElse Not isDefaultText Then Return

        Dim name = TxtName.Text.ToLower()
        Dim len = TxtName.Text.Length
        Dim suggestedName = ""

        If name.StartsWith("menuitem") Then
            suggestedName = TxtName.Text.Substring(8)
        ElseIf name.StartsWith("menu") Then
            suggestedName = TxtName.Text.Substring(4)
        ElseIf name.StartsWith("mnu") Then
            suggestedName = TxtName.Text.Substring(3)
        Else
            suggestedName = TxtName.Text
        End If

        If suggestedName <> "" Then TxtText.Text = suggestedName
    End Sub
End Class