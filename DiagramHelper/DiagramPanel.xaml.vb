Imports System.ComponentModel
Imports System.Globalization

Public Class DiagramPanel

    Friend ConnectorsGrid As Grid
    Friend FocusRectangle As Rectangle
    Friend Dsn As Designer
    Friend DesignerItem As ListBoxItem
    Friend Diagram As FrameworkElement
    Dim Scv As ScrollViewer
    Friend DiagramObj As DiagramObject
    Friend AfterRestoreSub As Action
    Friend DiagramGroup As DiagramGroup

    Public Overrides Sub OnApplyTemplate()
        MyBase.OnApplyTemplate()
        DesignerItem = Helper.GetListBoxItem(Me)

        Scv = Helper.GetScrollViewer(DesignerItem)
        Dsn = Helper.GetDesigner(Scv)

        DesignerItem.AllowDrop = True
        Me.AllowDrop = True

        Diagram = TryCast(CType(Me.Content, ContentPresenter).Content, FrameworkElement)
        If Diagram Is Nothing Then Return

        Panel.SetZIndex(DesignerItem, Panel.GetZIndex(Diagram))

        Diagram.AllowDrop = True
        Diagram.Cursor = Cursors.SizeAll

        Canvas.SetLeft(DesignerItem, Designer.GetLeft(Diagram))
        Canvas.SetRight(DesignerItem, Designer.GetRight(Diagram))
        Canvas.SetTop(DesignerItem, Designer.GetTop(Diagram))
        Canvas.SetBottom(DesignerItem, Designer.GetBottom(Diagram))

        Me.Width = Helper.FixToMm(Designer.GetFrameWidth(Diagram)) + Diagram.Margin.Left + Diagram.Margin.Right
        Me.Height = Helper.FixToMm(Designer.GetFrameHeight(Diagram)) + Diagram.Margin.Top + Diagram.Margin.Bottom

        Dim Angle = Designer.GetRotationAngle(Diagram)
        Helper.Rotate(DesignerItem, Angle)

        Dim Cntrl As Control = TryCast(Diagram, Control)
        If Cntrl IsNot Nothing Then Cntrl.IsTabStop = False

        FocusRectangle = Me.Template.FindName("PART_FocusRectangle", Me)

        AddHandler DesignerItem.PreviewKeyDown, AddressOf DesignerItem_PreviewKeyDown
        AddHandler DesignerItem.LostFocus, AddressOf DesignerItem_LostFocus
        AddHandler DesignerItem.GotFocus, AddressOf DesignerItem_GotFocus
    End Sub

    Private Sub DiagramPanel_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        DiagramObj = DiagramObject.CreateDiagramObject(Diagram)
        AfterRestoreSub = AddressOf DiagramObj.AfterRestoreAction

        Dim FontProps = Designer.GetDiagramTextFontProps(Diagram)
        If FontProps Is Nothing Then
            FontProps = New PropertyDictionary(Diagram)
            Designer.SetDiagramTextFontProps(Diagram, FontProps)
        Else
            FontProps.SetDependencyObject(Diagram)
            FontProps.SetValuesToObj()
        End If

        Dim Id = Designer.GetGroupID(Diagram)
        If Id > 0 Then DiagramGroup.Add(Me, Id)

        Dim sk = TryCast(Diagram.LayoutTransform, SkewTransform)
        If sk Is Nothing OrElse sk.AngleY = 0 Then Diagram.LayoutTransform = New SkewTransform(0, 0.000000000001, 0, 0.000000000001)

        TopDpd = DependencyPropertyDescriptor.FromProperty(Canvas.TopProperty, GetType(ListBoxItem))
        TopDpd.AddValueChanged(DesignerItem, AddressOf DesignerItem_TopChanged)

        LeftDpd = DependencyPropertyDescriptor.FromProperty(Canvas.LeftProperty, GetType(ListBoxItem))
        LeftDpd.AddValueChanged(DesignerItem, AddressOf DesignerItem_LeftChanged)

        RenderTransformDpd = DependencyPropertyDescriptor.FromProperty(ListBoxItem.RenderTransformProperty, GetType(ListBoxItem))
        RenderTransformDpd.AddValueChanged(DesignerItem, AddressOf DesignerItem_RenderTransformChanged)

        TextDpd = DependencyPropertyDescriptor.FromProperty(TextBlock.TextProperty, GetType(TextBlock))
        TextDpd.AddValueChanged(Diagram, AddressOf DiagramObj.DiagramTextBlock_TextChanged)

        'Dim x = DependencyPropertyDescriptor.FromProperty(TextBlock.FontWeightProperty, GetType(TextBlock))
        'x.AddValueChanged(DiagramTextBlock, AddressOf DiagramObj.DiagramTextBlock_FontWeightChanged)

    End Sub

    Public Property IsSelected As Boolean
        Get
            Return GetValue(IsSelectedProperty)
        End Get

        Set(value As Boolean)
            SetValue(IsSelectedProperty, value)
        End Set
    End Property

    Public Shared ReadOnly IsSelectedProperty As DependencyProperty =
                           DependencyProperty.Register("IsSelected",
                           GetType(Boolean), GetType(DiagramPanel), New PropertyMetadata(AddressOf OnIsSelectedChanged))

    Public Event IsSelectedChanged(Pnl As DiagramPanel, NewValue As Boolean)

    Sub OnIsSelectedChanged(NewValue As Boolean)
        RaiseEvent IsSelectedChanged(Me, NewValue)
    End Sub

    Friend ExitIsSelectedChanged As Boolean = False

    Shared Sub OnIsSelectedChanged(Pnl As DiagramPanel, e As DependencyPropertyChangedEventArgs)
        If Pnl.ExitIsSelectedChanged Then Return
        If e.NewValue = False Then Pnl.Dsn.LocationVisibility = Windows.Visibility.Collapsed
        Pnl.OnIsSelectedChanged(e.NewValue)
    End Sub

    Sub Dispose()
        DiagramGroup.RemovePanelOnly(Me)

        RemoveHandler DesignerItem.PreviewKeyDown, AddressOf DesignerItem_PreviewKeyDown
        RemoveHandler DesignerItem.LostFocus, AddressOf DesignerItem_LostFocus
        RemoveHandler DesignerItem.GotFocus, AddressOf DesignerItem_GotFocus

        TopDpd.RemoveValueChanged(DesignerItem, AddressOf DesignerItem_TopChanged)
        LeftDpd.RemoveValueChanged(DesignerItem, AddressOf DesignerItem_LeftChanged)
        RenderTransformDpd.RemoveValueChanged(DesignerItem, AddressOf DesignerItem_RenderTransformChanged)
        TextDpd.RemoveValueChanged(Diagram, AddressOf DiagramObj.DiagramTextBlock_TextChanged)

        Dim Cp As ContentPresenter = VisualTreeHelper.GetParent(Diagram)
        Cp.Content = Nothing
        ConnectorsGrid = Nothing
        FocusRectangle = Nothing
        DesignerItem = Nothing
        Diagram = Nothing
        Scv = Nothing
        DiagramObj = Nothing
        AfterRestoreSub = Nothing
        TopDpd = Nothing
        LeftDpd = Nothing
        RenderTransformDpd = Nothing
        TextDpd = Nothing
    End Sub

    Private Sub DiagramPanel_SizeChanged(sender As Object, e As SizeChangedEventArgs) Handles Me.SizeChanged
        Helper.UpdateControl(Me)
        Designer.SetFrameWidth(Diagram, Me.ActualWidth)
        Designer.SetFrameHeight(Diagram, Me.ActualHeight)
    End Sub

    Sub ApplyLastChangeToSelected()
        If Commands.Cancelled Then
            Commands.Cancelled = False
            Return
        End If

        Dim Unit = Dsn.UndoStack.Peek
        If Unit Is Nothing Then Return
        Dim n = Unit.Count - 1

        For Each fw As FrameworkElement In Dsn.SelectedItems
            If fw Is Diagram Then Continue For

            For i = 0 To n
                Dim propState = TryCast(Unit(i), PropertyState)
                If PropState Is Nothing Then Continue For

                Dim target = GetTarget(fw, propState.Owner)
                Dim NewPropState As New PropertyState(target, propState.Keys.ToArray)

                For Each pair In propState
                    Dim oldValue = pair.Value.OldValue
                    Dim newValue = pair.Value.NewValue
                    If oldValue Is Nothing OrElse (oldValue IsNot newValue AndAlso Not oldValue.Equals(newValue)) Then
                        target.SetValue(pair.Key, Helper.Clone(pair.Value.NewValue))
                    End If
                Next

                If NewPropState.HasChanges Then Unit.Add(NewPropState.SetNewValue())
            Next
        Next
    End Sub

    Private Shared Function GetTarget(fw As FrameworkElement, stateOwner As DependencyObject) As Object
        If TypeOf stateOwner Is ListBoxItem Then
            Return Helper.GetListBoxItem(fw)
        ElseIf TypeOf stateOwner Is DiagramPanel Then
            Return Helper.GetDiagramPanel(fw)
        Else
            Return fw
        End If
    End Function


#Region "DesignerItem"

    Private Sub DesignerItem_GotFocus(sender As Object, e As RoutedEventArgs)
        If Dsn.SelectionBorder?.Visibility = Windows.Visibility.Visible Then Return

        FocusRectangle.StrokeThickness = 2
        If Dsn.IsReloading Then
            Dsn.IsReloading = False
        Else
            Me.IsSelected = True
        End If

        Dsn.ScrollIntoView(Diagram)
    End Sub

    Public Sub BringToFront()
        Dsn.MaxZIndex += 1
        Dim OldState = New PropertyState(DesignerItem, Canvas.ZIndexProperty)
        Canvas.SetZIndex(DesignerItem, Dsn.MaxZIndex)
        Dsn.UndoStack.ReportChanges(New UndoRedoUnit(OldState.SetNewValue()))
    End Sub

    Public Sub SendToBack()
        Dsn.MinZIndex -= 1
        Dim OldState = New PropertyState(DesignerItem, Canvas.ZIndexProperty)
        Canvas.SetZIndex(DesignerItem, Dsn.MinZIndex)
        Dsn.UndoStack.ReportChanges(New UndoRedoUnit(OldState.SetNewValue()))
    End Sub

    Private Sub DesignerItem_LostFocus(sender As Object, e As RoutedEventArgs)
        If Not Designer.Editing Then FocusRectangle.StrokeThickness = 0
    End Sub

    Private Sub DesignerItem_PreviewKeyDown(sender As Object, e As KeyEventArgs)
        If DesignerItem Is Nothing Then Return

        If Not DesignerItem.IsSelected Then
            Dsn.Focus()
            Return
        End If

        Dim offset = If(Keyboard.Modifiers = ModifierKeys.Control, Helper.CmToPx, Helper.MmToPx)
        Dsn.SelectedBounds = Dsn.GetSelectionBounds

        Select Case e.Key
            Case Key.F2 ' Rename
                DiagramObj.BeginEdit(True)
                e.Handled = True
            Case Key.F9
                Commands.ChangeBackground(Diagram)
                ApplyLastChangeToSelected()
                e.Handled = True
            Case Key.F3
                Dim B = Commands.ChangeBorderBrush(Diagram)
                If B IsNot Nothing Then ApplyLastChangeToSelected()
                e.Handled = True
            Case Key.F4
                Commands.ApplyLastChangeTo(Diagram)
                ApplyLastChangeToSelected()
                e.Handled = True
            Case Key.F6
                If Keyboard.Modifiers = ModifierKeys.Shift Then
                    Commands.IncreaseBorderThickness(Diagram, -0.1)
                Else
                    Commands.IncreaseBorderThickness(Diagram, 0.1)
                End If

                ApplyLastChangeToSelected()
                e.Handled = True
            Case Key.F7
                Commands.ChangeBrush(Diagram, Control.ForegroundProperty)
                ApplyLastChangeToSelected()
            Case Key.F11
                Commands.IncreaseRotationAngle(Diagram, -45)
                ApplyLastChangeToSelected()
            Case Key.F12
                Commands.IncreaseRotationAngle(Diagram, 45)
                ApplyLastChangeToSelected()
            Case Key.K
                If Keyboard.Modifiers = ModifierKeys.Control Then
                    Commands.Skew(Diagram)
                    ApplyLastChangeToSelected()
                End If
            Case Key.F
                If Keyboard.Modifiers = ModifierKeys.Control Then
                    Commands.ChangeFont(Diagram)
                    ApplyLastChangeToSelected()
                End If
            Case Key.Oem4
                If Keyboard.Modifiers = ModifierKeys.Control Then
                    Commands.IncreaseFontSize(Diagram, -1)
                    ApplyLastChangeToSelected()
                End If
            Case Key.Oem6
                If Keyboard.Modifiers = ModifierKeys.Control Then
                    Commands.IncreaseFontSize(Diagram, +1)
                    ApplyLastChangeToSelected()
                End If
            Case Key.B
                If Keyboard.Modifiers = ModifierKeys.Control Then
                    MenuItem_SubmenuOpened(Nothing, Nothing)
                    BoldMenuItem.IsChecked = Not BoldMenuItem.IsChecked
                End If
            Case Key.I
                If Keyboard.Modifiers = ModifierKeys.Control Then
                    MenuItem_SubmenuOpened(Nothing, Nothing)
                    ItalicMenuItem.IsChecked = Not ItalicMenuItem.IsChecked
                End If
            Case Key.U
                If Keyboard.Modifiers = ModifierKeys.Control Then
                    MenuItem_SubmenuOpened(Nothing, Nothing)
                    UnderlineMenuItem.IsChecked = Not UnderlineMenuItem.IsChecked
                End If
            Case Key.G
                If Keyboard.Modifiers = ModifierKeys.Control Then
                    GroupMenuItem.IsChecked = Not GroupMenuItem.IsChecked
                End If
            Case Key.Enter
                DiagramObj.BeginEdit()
                e.Handled = True
            Case Key.Delete
                Dsn.RemoveSelectedItems()
            Case Key.Tab
                If Keyboard.Modifiers = ModifierKeys.Shift Then
                    Dim I = Dsn.Items.IndexOf(Diagram) - 1
                    If I > -1 Then
                        Dim Itm = Helper.GetListBoxItem(Dsn.Items(I))
                        If Not Itm.IsSelected Then Dsn.SelectedIndex = I
                        Itm.Focus()
                    Else
                        Dsn.MoveFocus(New TraversalRequest(FocusNavigationDirection.Previous))
                    End If
                Else
                    Dim I = Dsn.Items.IndexOf(Diagram) + 1
                    If I < Dsn.Items.Count Then
                        Dim Itm = Helper.GetListBoxItem(Dsn.Items(I))
                        If Not Itm.IsSelected Then Dsn.SelectedIndex = I
                        Itm.Focus()
                    Else
                        Dim g As Grid = Dsn.Parent
                        Dsn.IsEnabled = False
                        Dsn.MoveFocus(New TraversalRequest(FocusNavigationDirection.Next))
                        Dsn.IsEnabled = True
                    End If
                End If
                e.Handled = True
            Case Key.Up
                DiagramObj.MoveDiagram(0, -offset, True)
                e.Handled = True
            Case Key.PageUp
                If Keyboard.Modifiers = ModifierKeys.Control Then
                    DiagramObj.MoveDiagram(-Scv.ViewportWidth, 0, True)
                Else
                    DiagramObj.MoveDiagram(0, -Scv.ViewportHeight, True)
                End If
                e.Handled = True
            Case Key.Left
                DiagramObj.MoveDiagram(-offset, 0, True)
                e.Handled = True
            Case Key.Down
                DiagramObj.MoveDiagram(0, offset, True)
                e.Handled = True
            Case Key.PageDown
                If Keyboard.Modifiers = ModifierKeys.Control Then
                    DiagramObj.MoveDiagram(Scv.ViewportWidth, 0, True)
                Else
                    DiagramObj.MoveDiagram(0, Scv.ViewportHeight, True)
                End If
                e.Handled = True
            Case Key.Right
                DiagramObj.MoveDiagram(offset, 0, True)
                e.Handled = True
        End Select
    End Sub

    Dim TopDpd As DependencyPropertyDescriptor
    Dim LeftDpd As DependencyPropertyDescriptor
    Dim RenderTransformDpd As DependencyPropertyDescriptor
    Dim TextDpd As DependencyPropertyDescriptor

    Private Sub DesignerItem_TopChanged()
        If DesignerItem Is Nothing Then Return
        Designer.SetTop(Diagram, Canvas.GetTop(DesignerItem))
    End Sub

    Private Sub DesignerItem_LeftChanged()
        If DesignerItem Is Nothing Then Return
        Designer.SetLeft(Diagram, Canvas.GetLeft(DesignerItem))
    End Sub

    Private Sub DesignerItem_RenderTransformChanged()
        If DesignerItem Is Nothing Then Return

        Dim RotateTransform = TryCast(DesignerItem.RenderTransform, RotateTransform)
        If RotateTransform Is Nothing Then
            Designer.SetRotationAngle(Diagram, 0)
        Else
            Designer.SetRotationAngle(Diagram, RotateTransform.Angle)
        End If
    End Sub

#End Region

#Region "Menus"

    Private Sub EditTextMenuItem_Click(sender As Object, e As RoutedEventArgs)
        DiagramObj.BeginEdit()
    End Sub

    Private Sub DiagramBackgroundMenuItem_Click(sender As Object, e As RoutedEventArgs)
        Commands.ChangeBackground(Diagram)
        ApplyLastChangeToSelected()
    End Sub

    Private Sub DiagramBorderBrushMenuItem_Click(sender As Object, e As RoutedEventArgs)
        Dim B = Commands.ChangeBorderBrush(Diagram)
        If B IsNot Nothing Then ApplyLastChangeToSelected()
    End Sub

    Private Sub IicreaseBorderThicknessMenuItem_Click(sender As Object, e As RoutedEventArgs)
        Commands.IncreaseBorderThickness(Diagram, 0.1)
        ApplyLastChangeToSelected()
    End Sub

    Private Sub DecreaseBorderThicknessMenuItem_Click(sender As Object, e As RoutedEventArgs)
        Commands.IncreaseBorderThickness(Diagram, -0.1)
        ApplyLastChangeToSelected()
    End Sub

    Private Sub TextForegroundMenuItem_Click(sender As Object, e As RoutedEventArgs)
        Commands.ChangeBrush(Diagram, Control.ForegroundProperty)
        ApplyLastChangeToSelected()
    End Sub

    Private Sub ZeroRotationMenuItem_Click(sender As Object, e As RoutedEventArgs)
        Commands.IncreaseRotationAngle(Diagram, -Designer.GetRotationAngle(Diagram))
        ApplyLastChangeToSelected()
    End Sub

    Private Sub DecreaseRotatationMenuItem_Click(sender As Object, e As RoutedEventArgs)
        Commands.IncreaseRotationAngle(Diagram, -45)
        ApplyLastChangeToSelected()
    End Sub

    Private Sub IncreaseRotatationMenuItem_Click(sender As Object, e As RoutedEventArgs)
        Commands.IncreaseRotationAngle(Diagram, 45)
        ApplyLastChangeToSelected()
    End Sub


    Private Sub ApplyLastChangeMenuItem_Click(sender As Object, e As RoutedEventArgs)
        Commands.ApplyLastChangeTo(Diagram)
        ApplyLastChangeToSelected()
    End Sub

    Private Sub SkewMenuItem_Click(sender As Object, e As RoutedEventArgs)
        Commands.Skew(Diagram)
        ApplyLastChangeToSelected()
    End Sub

    Private Sub RenameMenuItem_Click(sender As Object, e As RoutedEventArgs)
        DiagramObj.BeginEdit(True)
    End Sub

    Private Sub CopyMenuItem_Click(sender As Object, e As RoutedEventArgs)
        Dsn.Copy()
    End Sub

    Private Sub CutMenuItem_Click(sender As Object, e As RoutedEventArgs)
        Dsn.Cut()
    End Sub

    Private Sub DeleteMenuItem_Click(sender As Object, e As RoutedEventArgs)
        Dsn.RemoveSelectedItems()
    End Sub

    Private Sub TextFontMenuItem_Click(sender As Object, e As RoutedEventArgs)
        Commands.ChangeFont(Diagram)
        ApplyLastChangeToSelected()
    End Sub

    Private Sub DecreaseSizeMenuItem_Click(sender As Object, e As RoutedEventArgs)
        Commands.IncreaseFontSize(Diagram, -1)
        ApplyLastChangeToSelected()
    End Sub

    Private Sub IncreaseSizeMenuItem_Click(sender As Object, e As RoutedEventArgs)
        Commands.IncreaseFontSize(Diagram, 1)
        ApplyLastChangeToSelected()
    End Sub

    Private Sub BoldMenuItem_Checked(sender As Object, e As RoutedEventArgs)
        If freezMenuCheck Then Return

        If BoldMenuItem.IsChecked Then
            Commands.ChangeFontProperty(Diagram, TextBlock.FontWeightProperty, FontWeights.Bold)
        Else
            Commands.ChangeFontProperty(Diagram, TextBlock.FontWeightProperty, FontWeights.Normal)
        End If
        ApplyLastChangeToSelected()
    End Sub

    Private Sub ItalicMenuItem_Checked(sender As Object, e As RoutedEventArgs)
        If freezMenuCheck Then Return

        If ItalicMenuItem.IsChecked Then
            Commands.ChangeFontProperty(Diagram, TextBlock.FontStyleProperty, FontStyles.Italic)
        Else
            Commands.ChangeFontProperty(Diagram, TextBlock.FontStyleProperty, FontStyles.Normal)
        End If
        ApplyLastChangeToSelected()
    End Sub

    Private Sub UnderlineMenuItem_Checked(sender As Object, e As RoutedEventArgs)
        If freezMenuCheck Then Return

        Dim Decorations As TextDecorationCollection = Nothing
        If UnderlineMenuItem.IsChecked Then
            Decorations = TextDecorations.Underline
        End If

        Commands.ChangeFontProperty(
            GetTextObject(),
            TextBlock.TextDecorationsProperty,
            Decorations
        )
        ApplyLastChangeToSelected()
    End Sub

    Private Function GetTextObject() As FrameworkElement
        Dim contentControl = TryCast(Diagram, ContentControl)
        Dim obj = Nothing

        If contentControl IsNot Nothing Then
            obj = contentControl.Content
            If TypeOf obj Is String Then
                obj = New TextBlock With {
                    .Text = obj,
                    .TextWrapping = TextWrapping.Wrap
                }
                contentControl.Content = obj
            End If
            Return obj
        End If

        Return Diagram
    End Function


#End Region

    Private Sub ContextMenu_Opened(sender As Object, e As RoutedEventArgs)
        If Me.DiagramGroup Is Nothing Then
            GroupMenuItem.IsEnabled = (Dsn.SelectedItems.Count > 1)
            RemoveFromGroupMenuItem.Visibility = Visibility.Collapsed
        Else
            GroupMenuItem.IsEnabled = True
            If DiagramGroup.Count > 2 Then
                RemoveFromGroupMenuItem.Visibility = Visibility.Visible
            Else
                RemoveFromGroupMenuItem.Visibility = Visibility.Collapsed
            End If
        End If
    End Sub

    Friend ExitGroupChecked As Boolean = False

    Private Sub GroupMenuItem_Checked(sender As Object, e As RoutedEventArgs)
        If ExitGroupChecked Then Return

        Dim UndoUnit As New UndoRedoUnit
        Dim timeStamp = Now.Ticks

        For Each d As FrameworkElement In Dsn.SelectedItems
            Dim act As Action = AddressOf DiagramObject.Diagrams(d).AfterRestoreAction
            Dim OldSate As New PropertyState(act, d, Designer.GroupIDProperty)
            Designer.SetGroupID(d, timeStamp)
            UndoUnit.Add(OldSate.SetNewValue)
        Next

        Dsn.UndoStack.ReportChanges(UndoUnit)
        DiagramGroup.Select()
    End Sub

    Private Sub GroupMenuItem_Unchecked(sender As Object, e As RoutedEventArgs)
        If ExitGroupChecked Then Return
        DiagramGroup.Ungroup()
    End Sub

    Private Sub RemoveFromGroupMenuItem_Click(sender As Object, e As RoutedEventArgs)
        Diagram.ClearValue(Designer.GroupIDProperty)
    End Sub


    Friend Sub UpdateLocationBorder()
        FocusRectangle.LayoutTransform = Diagram.LayoutTransform
        FocusRectangle.Width = Diagram.ActualWidth
        FocusRectangle.Height = Diagram.ActualHeight

        FocusRectangle.StrokeThickness = 2
        Dsn.LocationVisibility = Windows.Visibility.Visible
        UpdateTbLocation()
    End Sub

    Sub UpdateTbLocation()
        Dim P As Point = DiagramObj.GetLeftTopPoint()
        Dsn.TbLeftLocation.Text = P.X.ToString("F1")
        Dsn.TbTopLocation.Text = P.Y.ToString("F1")
        Helper.UpdateControl(Dsn.TbLeftLocation)
        Helper.UpdateControl(Dsn.TbTopLocation)

        'Dim Location = GetLeftTopPoint(Dsn, False)
        Dim x = P.X * Helper.CmToPx * Dsn.Scale - 5
        Dim y = P.Y * Helper.CmToPx * Dsn.Scale - 5
        Dsn.TbLeftLocation.Margin = New Thickness(x - Dsn.TbLeftLocation.ActualWidth * Dsn.Scale, y + 2, 0, 0)
        Dsn.TbTopLocation.Margin = New Thickness(x + 2, y - Dsn.TbTopLocation.ActualHeight * Dsn.Scale, 0, 0)

    End Sub

    Private Sub BringToFrontMenuItem_Click(sender As Object, e As RoutedEventArgs)
        BringToFront()
    End Sub

    Private Sub SendToBackMenuItem_Click(sender As Object, e As RoutedEventArgs)
        SendToBack()
    End Sub

    Dim freezMenuCheck As Boolean

    Private Sub MenuItem_SubmenuOpened(sender As Object, e As RoutedEventArgs)
        Dim obj As Object = GetTextObject()
        freezMenuCheck = True

        If obj Is Nothing Then
            BoldMenuItem.IsChecked = False
            ItalicMenuItem.IsChecked = False
            UnderlineMenuItem.IsChecked = False

        Else
            Try
                BoldMenuItem.IsChecked = (obj.FontWeight = FontWeights.Bold)
            Catch
            End Try

            Try
                ItalicMenuItem.IsChecked = (obj.FontStyle = FontStyles.Italic)
            Catch
            End Try

            Try
                UnderlineMenuItem.IsChecked = (obj.TextDecorations Is TextDecorations.Underline)
            Catch
            End Try
        End If

        freezMenuCheck = False
    End Sub

    Private Sub PropertiesMenuItem_Click(sender As Object, e As RoutedEventArgs)
        Dim WndProps As New WndProperties
        With WndProps
            .Show()
            .Hide()
            .LeftValue = GetDoubleValue(DesignerItem, Canvas.LeftProperty)
            .TopValue = GetDoubleValue(DesignerItem, Canvas.TopProperty)
            .WidthValue = GetDoubleValue(Me, WidthProperty)
            .HeightValue = GetDoubleValue(Me, HeightProperty)
            .MaxWidthValue = GetDoubleValue(Me, MaxWidthProperty)
            .MaxHeightValue = GetDoubleValue(Me, MaxHeightProperty)
            .EnabledValue = Diagram.IsEnabled
            .VisibleValue = Diagram.IsVisible
            .RightToLeftValue = Diagram.FlowDirection = FlowDirection.RightToLeft
            .TagValue = Diagram.Tag
            .ToolTipValue = Diagram.ToolTip
        End With

        If WndProps.ShowDialog = True Then
            Dim unit As New UndoRedoUnit
            Dim muliItems = Dsn.SelectedItems.Count > 1

            Dim OldState As New PropertyState(DesignerItem, Canvas.LeftProperty, Canvas.TopProperty)
            If WndProps.LeftValue.HasValue Then Canvas.SetLeft(DesignerItem, WndProps.LeftValue)
            If WndProps.TopValue.HasValue Then Canvas.SetTop(DesignerItem, WndProps.TopValue)
            If muliItems OrElse OldState.HasChanges Then unit.Add(OldState.SetNewValue)

            OldState = New PropertyState(Me, FrameworkElement.WidthProperty, FrameworkElement.HeightProperty)
            If WndProps.WidthValue.HasValue Then Me.Width = WndProps.WidthValue
            If WndProps.HeightValue.HasValue Then Me.Height = WndProps.HeightValue
            If muliItems OrElse OldState.HasChanges Then unit.Add(OldState.SetNewValue)

            OldState = New PropertyState(Me, FrameworkElement.MaxWidthProperty, FrameworkElement.MaxHeightProperty)
            If WndProps.MaxWidthValue.HasValue Then Me.MaxWidth = WndProps.MaxWidthValue
            If WndProps.MaxHeightValue.HasValue Then Me.MaxHeight = WndProps.MaxHeightValue
            If muliItems OrElse OldState.HasChanges Then unit.Add(OldState.SetNewValue)

            If unit.Count > 0 Then
                Dsn.UndoStack.ReportChanges(unit)
                ApplyLastChangeToSelected()
            End If
        End If
    End Sub

    Private Function GetDoubleValue(item As FrameworkElement, dp As DependencyProperty) As Double?
        Dim value As Double = item.GetValue(dp)
        If Dsn.SelectedItems.Count = 1 Then Return value

        Dim onCanvus = TypeOf item Is ListBoxItem
        For Each fw As FrameworkElement In Dsn.SelectedItems
            If fw Is item Then Continue For

            If onCanvus Then
                If Not Helper.GetListBoxItem(fw).GetValue(dp).Equals(value) Then Return Nothing
            ElseIf Not Helper.GetDiagramPanel(fw).GetValue(dp).Equals(value) Then
                Return Nothing
            End If
        Next

        Return value
    End Function
End Class
