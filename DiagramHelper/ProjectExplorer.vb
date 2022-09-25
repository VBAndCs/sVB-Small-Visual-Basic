Imports System.Windows.Controls.Primitives
Imports System.ComponentModel
Imports System.Windows.Threading

Public Class ProjectExplorer
    Inherits Control

    Public Sub New()
        Dim resourceLocater As Uri = New Uri("/DiagramHelper;component/Resources/projectexplorerdecorator.xaml", System.UriKind.Relative)
        Dim ResDec As ResourceDictionary = Application.LoadComponent(resourceLocater)
        Me.Resources.MergedDictionaries.Add(ResDec)
        Me.Style = FindResource("ProjectExplorerStyle")
    End Sub

    Public WithEvents FilesList As ListBox

    Dim FirstTime As Boolean = True
    Public Event ItemDoubleClick(sender As Object, e As MouseButtonEventArgs)

    Private Sub ProjectExplorer_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        If FirstTime Then
            FirstTime = False

            FilesList = TryCast(Template.FindName("PART_ListBox", Me), ListBox)
            FilesList.ItemContainerStyle = FindResource("listBoxItemStyle")
            FilesList.ItemsSource = Designer.FormNames
            FilesList.SelectedIndex = 0
        End If
    End Sub

    Public FreezListFiles As Boolean

    Dim SelectedAt As Date

    Sub FilesList_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles FilesList.SelectionChanged
        If FreezListFiles Then Return
        Dim i = FilesList.SelectedIndex
        If i = -1 Then Return
        SelectedAt = Now
        Designer.SwitchTo(Designer.FormKeys(i))
    End Sub

    Private Sub FilesList_KeyDown(sender As Object, e As KeyEventArgs) Handles FilesList.KeyDown
        Select Case e.Key
            Case Key.F2
                BeginEdit()
                e.Handled = True

            Case Key.Enter
                RaiseEvent ItemDoubleClick(Me, Nothing)

            Case Key.Delete
                If Designer.IsNew AndAlso Designer.Pages.Count = 1 Then
                    Beep()
                    Return
                End If

                If FilesList.SelectedIndex = -1 Then Return
                Designer.ClosePage()
        End Select
    End Sub


#Region "Edit Item Name"
    Dim item As ListBoxItem
    Dim txtBlock As TextBlock
    Dim editorGrid As Grid
    Dim editTextBox As TextBox
    Dim inEditMode As Boolean = False

    Dim ClickCount As Integer
    Private Sub FilesList_PreviewMouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs) Handles FilesList.PreviewMouseLeftButtonDown
        item = CType(Helper.GetListBoxItem(e.OriginalSource), ListBoxItem)
        If item Is Nothing Then
            If inEditMode Then Commit()

        ElseIf e.ClickCount = 2 Then
            ClickCount = 2
            If Helper.GetParent(Of TextBox)(e.OriginalSource) IsNot Nothing Then Return
            RaiseEvent ItemDoubleClick(sender, e)

        ElseIf (Now - SelectedAt).TotalMilliseconds > 200 AndAlso
                    TypeOf e.OriginalSource Is TextBlock AndAlso
                    item.IsSelected AndAlso item.IsFocused Then

            ClickCount = 1
            RunAction.After(200,
                   Sub()
                       If ClickCount = 2 Then Return
                       BeginEdit()
                       e.Handled = True
                       SelectedAt = Now
                   End Sub)
        End If
    End Sub

    Public Sub BeginEdit()
        inEditMode = True
        Dim brdr As Border = VisualTreeHelper.GetChild(item, 0)
        Dim grid As Grid = brdr.Child
        txtBlock = grid.Children(0)
        editorGrid = grid.Children(1)
        editTextBox = editorGrid.Children(1)

        txtBlock.Visibility = Visibility.Collapsed
        ' don't use textBlock because it is bound to data context
        Dim s = CStr(item.Content).Substring(3).Trim(" ", "*")
        If s.EndsWith(".xaml") Then s = s.Substring(0, s.Length - 5)
        editTextBox.Text = s
        editorGrid.Visibility = Visibility.Visible
        editTextBox.Focus()
        editTextBox.SelectAll()

        AddHandler editTextBox.PreviewKeyDown, AddressOf editTextBox_PreviewKeyDown
        AddHandler editTextBox.PreviewTextInput, AddressOf editTextBox_PreviewTextInput
        AddHandler editTextBox.PreviewLostKeyboardFocus, AddressOf editTextBox_LostFocus
    End Sub

    Public Sub CancelEdit()
        RemoveHandler editTextBox.PreviewKeyDown, AddressOf editTextBox_PreviewKeyDown
        RemoveHandler editTextBox.PreviewTextInput, AddressOf editTextBox_PreviewTextInput
        RemoveHandler editTextBox.PreviewLostKeyboardFocus, AddressOf editTextBox_LostFocus

        editorGrid.Visibility = Visibility.Collapsed
        txtBlock.Visibility = Visibility.Visible
        inEditMode = False
    End Sub

    Public Function Commit() As Boolean
        Dim newName = editTextBox.Text
        If Designer.ChangeFormName(newName) Then
            editorGrid.Visibility = Visibility.Collapsed
            txtBlock.Visibility = Visibility.Visible
            SelectedAt = Now
            inEditMode = False
            Return True
        End If

        Return False
    End Function

    Private Sub editTextBox_PreviewKeyDown(sender As Object, e As KeyEventArgs)
        Select Case e.Key
            Case Key.Space
                Beep()
                e.Handled = True

            Case Key.Escape
                CancelEdit()
                e.Handled = True
                item.Focus()

            Case Key.Enter
                If Commit() Then
                    RemoveHandler editTextBox.PreviewKeyDown, AddressOf editTextBox_PreviewKeyDown
                    RemoveHandler editTextBox.PreviewTextInput, AddressOf editTextBox_PreviewTextInput
                    RemoveHandler editTextBox.PreviewLostKeyboardFocus, AddressOf editTextBox_LostFocus
                    item.Focus()
                End If
                e.Handled = True

        End Select
    End Sub

    Public Sub editTextBox_PreviewTextInput(sender As Object, e As TextCompositionEventArgs)
        Select Case e.Text.ToLower
            Case "a" To "z", "_", "0" To "9"
                ' allowed
            Case Else
                Beep()
                e.Handled = True
        End Select
    End Sub


    Dim focusTextBox As New RunAction(10, Sub() editTextBox.Focus())

    Public Sub editTextBox_LostFocus(sender As Object, e As KeyboardFocusChangedEventArgs)
        If Commit() Then
            RemoveHandler editTextBox.PreviewKeyDown, AddressOf editTextBox_PreviewKeyDown
            RemoveHandler editTextBox.PreviewTextInput, AddressOf editTextBox_PreviewTextInput
            RemoveHandler editTextBox.PreviewLostKeyboardFocus, AddressOf editTextBox_LostFocus
        Else
            e.Handled = True
            focusTextBox.Start()
        End If
    End Sub



#End Region

    Public Property Designer As Designer
        Get
            Return GetValue(DesignerProperty)
        End Get

        Set(value As Designer)
            SetValue(DesignerProperty, value)
        End Set
    End Property

    Public Shared ReadOnly DesignerProperty As DependencyProperty =
                           DependencyProperty.Register("Designer",
                           GetType(Designer), GetType(ProjectExplorer))

End Class
