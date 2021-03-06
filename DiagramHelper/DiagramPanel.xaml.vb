﻿Imports System.ComponentModel
Imports System.Globalization

Public Class DiagramPanel

    Friend ConnectorsGrid As Grid
    Friend FocusRectangle As Rectangle
    Friend DiagramTextBlock As TextBlock
    Friend MyDesigner As Designer
    Friend DesignerItem As ListBoxItem
    Friend Diagram As FrameworkElement
    Dim EditorShowing As Boolean
    Dim Scv As ScrollViewer
    Friend DiagramObj As DiagramObject
    Friend AfterRestoreSub As Action
    Friend DiagramGroup As DiagramGroup

    Friend Event ConnectorsPositionChangd()

    Friend Sub OnConnectorsPositionChangd()
        Helper.UpdateControl(Me)
        RaiseEvent ConnectorsPositionChangd()
    End Sub


    Public Overrides Sub OnApplyTemplate()
        MyBase.OnApplyTemplate()
        DesignerItem = Helper.GetListBoxItem(Me)

        Scv = Helper.GetScrollViewer(DesignerItem)
        MyDesigner = Helper.GetDesigner(Scv)

        DesignerItem.AllowDrop = True
        Me.AllowDrop = True

        Diagram = TryCast(CType(Me.Content, ContentPresenter).Content, FrameworkElement)
        If Diagram Is Nothing Then Return

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

        DiagramTextBlock = Me.Template.FindName("PART_Text", Me)
        ConnectorsGrid = DiagramTextBlock.Parent
        Designer.SetDiagramTextBlock(Diagram, DiagramTextBlock)

        AddHandler DesignerItem.PreviewKeyDown, AddressOf DesignerItem_PreviewKeyDown
        AddHandler DesignerItem.LostFocus, AddressOf DesignerItem_LostFocus
        AddHandler DesignerItem.GotFocus, AddressOf DesignerItem_GotFocus
    End Sub

    Private Sub DiagramPanel_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        DiagramObj = DiagramObject.CreateDiagramObject(Diagram)
        AfterRestoreSub = AddressOf DiagramObj.AfterRestoreAction

        DiagramTextBlock.Text = Designer.GetDiagramText(Diagram)
        DiagramTextBlock.Background = Designer.GetDiagramTextBackground(Diagram)
        DiagramTextBlock.Foreground = Designer.GetDiagramTextForeground(Diagram)

        Dim FontProps = Designer.GetDiagramTextFontProps(Diagram)
        If FontProps Is Nothing Then
            FontProps = New PropertyDictionary(DiagramTextBlock)
            Designer.SetDiagramTextFontProps(Diagram, FontProps)
        Else
            FontProps.SetDependencyObject(DiagramTextBlock)
            FontProps.SetValuesToObj()
        End If

        ApplyRotationMenuItem.IsChecked = Designer.GetDiagramTextApplyRotation(Diagram)
        DrawOutlineMenuItem.IsChecked = Designer.GetDiagramTextOutlined(Diagram)

        Dim Id = Designer.GetGroupID(Diagram)
        If Id <> "" Then DiagramGroup.Add(Me, Date.Parse(Id))

        Dim sk = TryCast(Diagram.LayoutTransform, SkewTransform)
        If sk Is Nothing OrElse sk.AngleY = 0 Then Diagram.LayoutTransform = New SkewTransform(0, 0.000000000001, 0, 0.000000000001)

        TopDpd = DependencyPropertyDescriptor.FromProperty(Canvas.TopProperty, GetType(ListBoxItem))
        TopDpd.AddValueChanged(DesignerItem, AddressOf DesignerItem_TopChanged)

        LeftDpd = DependencyPropertyDescriptor.FromProperty(Canvas.LeftProperty, GetType(ListBoxItem))
        LeftDpd.AddValueChanged(DesignerItem, AddressOf DesignerItem_LeftChanged)

        RenderTransformDpd = DependencyPropertyDescriptor.FromProperty(ListBoxItem.RenderTransformProperty, GetType(ListBoxItem))
        RenderTransformDpd.AddValueChanged(DesignerItem, AddressOf DesignerItem_RenderTransformChanged)

        TextDpd = DependencyPropertyDescriptor.FromProperty(TextBlock.TextProperty, GetType(TextBlock))
        TextDpd.AddValueChanged(DiagramTextBlock, AddressOf DiagramObj.DiagramTextBlock_TextChanged)

        'Dim x = DependencyPropertyDescriptor.FromProperty(TextBlock.FontWeightProperty, GetType(TextBlock))
        'x.AddValueChanged(DiagramTextBlock, AddressOf DiagramObj.DiagramTextBlock_FontWeightChanged)

    End Sub

    Friend Sub AdjustConnectors()
        If Diagram Is Nothing Then Return
        FocusRectangle.LayoutTransform = Diagram.LayoutTransform
        FocusRectangle.Width = Diagram.ActualWidth
        FocusRectangle.Height = Diagram.ActualHeight
        If FocusRectangle.LayoutTransform IsNot Nothing Then
            Dim T = FocusRectangle.LayoutTransform.Inverse
            DiagramTextBlock.LayoutTransform = T
            For Each C In ConnectorsGrid.Children
                If TypeOf C Is ConnectorThumb Then
                    C.RenderTransformOrigin = New Point(0, 0)
                    C.LayoutTransform = T
                End If
            Next
        End If
    End Sub

    Public Property IsSelected As Boolean
        Get
            Return GetValue(IsSelectedProperty)
        End Get

        Set(ByVal value As Boolean)
            SetValue(IsSelectedProperty, value)
        End Set
    End Property

    Public Shared ReadOnly IsSelectedProperty As DependencyProperty = _
                           DependencyProperty.Register("IsSelected", _
                           GetType(Boolean), GetType(DiagramPanel), New PropertyMetadata(AddressOf OnIsSelectedChanged))

    Public Event IsSelectedChanged(Pnl As DiagramPanel, NewValue As Boolean)

    Sub OnIsSelectedChanged(NewValue As Boolean)
        RaiseEvent IsSelectedChanged(Me, NewValue)
    End Sub

    Friend ExitIsSelectedChanged As Boolean = False

    Shared Sub OnIsSelectedChanged(Pnl As DiagramPanel, e As DependencyPropertyChangedEventArgs)
        If Pnl.ExitIsSelectedChanged Then Return
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
        TextDpd.RemoveValueChanged(DiagramTextBlock, AddressOf DiagramObj.DiagramTextBlock_TextChanged)

        Dim Cp As ContentPresenter = VisualTreeHelper.GetParent(Diagram)
        Cp.Content = Nothing
        ConnectorsGrid = Nothing
        FocusRectangle = Nothing
        DiagramTextBlock = Nothing
        DesignerItem = Nothing
        Diagram = Nothing
        EditorShowing = Nothing
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
        Dim Unit = MyDesigner.UndoStack.Peek
        If Unit Is Nothing Then Return
        Dim PropState = TryCast(Unit(0), PropertyState)
        If PropState Is Nothing Then Return

        For Each Obj As FrameworkElement In MyDesigner.SelectedItems
            If Obj Is Diagram Then Continue For
            Dim NewPropState As New PropertyState(Obj, PropState.Keys.ToArray)
            For Each Pair In PropState
                Obj.SetValue(Pair.Key, Helper.Clone(Pair.Value.NewValue))
            Next
            If NewPropState.HasChanges Then Unit.Add(NewPropState.SetNewValue())
        Next
    End Sub


#Region "DesignerItem"

    Private Sub DesignerItem_GotFocus(sender As Object, e As RoutedEventArgs)
        If MyDesigner.SelectionBorder.Visibility = Windows.Visibility.Visible Then Return

        FocusRectangle.StrokeThickness = 2
        Me.IsSelected = True
        MyDesigner.MaxZIndex += 1
        Canvas.SetZIndex(DesignerItem, MyDesigner.MaxZIndex)
        MyDesigner.ScrollIntoView(Diagram)
    End Sub

    Private Sub DesignerItem_LostFocus(sender As Object, e As RoutedEventArgs)
        If Not Designer.Editing Then FocusRectangle.StrokeThickness = 0
    End Sub

    Private Sub DesignerItem_PreviewKeyDown(sender As Object, e As KeyEventArgs)
        If DesignerItem Is Nothing Then Return

        If Not DesignerItem.IsSelected Then
            MyDesigner.Focus()
            Return
        End If

        Dim offset = If(Keyboard.Modifiers = ModifierKeys.Control, Helper.CmToPx, Helper.MmToPx)
        MyDesigner.SelectedBounds = MyDesigner.GetSelectionBounds

        Select Case e.Key
            Case Key.F2
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
            Case Key.F5
                Commands.IncreaseBorderThickness(Diagram, -0.1)
                ApplyLastChangeToSelected()
                e.Handled = True
            Case Key.F6
                Commands.IncreaseBorderThickness(Diagram, 0.1)
                ApplyLastChangeToSelected()
                e.Handled = True
            Case Key.F7
                Commands.ChangeBrush(Diagram, Designer.DiagramTextForegroundProperty)
                ApplyLastChangeToSelected()
            Case Key.F8
                Commands.ChangeBrush(Diagram, Designer.DiagramTextBackgroundProperty)
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
            Case Key.OemPlus
                Commands.IncreaseTextOutlineThickness(Diagram, 0.1)
                ApplyLastChangeToSelected()
            Case Key.OemMinus
                Commands.IncreaseTextOutlineThickness(Diagram, -0.1)
                ApplyLastChangeToSelected()
            Case Key.B
                If Keyboard.Modifiers = ModifierKeys.Control Then
                    BoldMenuItem.IsChecked = Not BoldMenuItem.IsChecked
                End If
            Case Key.I
                If Keyboard.Modifiers = ModifierKeys.Control Then
                    ItalicMenuItem.IsChecked = Not ItalicMenuItem.IsChecked
                End If
            Case Key.U
                If Keyboard.Modifiers = ModifierKeys.Control Then
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
                MyDesigner.RemoveSelectedItems()
            Case Key.Tab
                If Keyboard.Modifiers = ModifierKeys.Shift Then
                    Dim I = MyDesigner.Items.IndexOf(Diagram) - 1
                    If I > -1 Then
                        Dim Itm = Helper.GetListBoxItem(MyDesigner.Items(I))
                        If Not Itm.IsSelected Then MyDesigner.SelectedIndex = I
                        Itm.Focus()
                    Else
                        MyDesigner.MoveFocus(New TraversalRequest(FocusNavigationDirection.Previous))
                    End If
                Else
                    Dim I = MyDesigner.Items.IndexOf(Diagram) + 1
                    If I < MyDesigner.Items.Count Then
                        Dim Itm = Helper.GetListBoxItem(MyDesigner.Items(I))
                        If Not Itm.IsSelected Then MyDesigner.SelectedIndex = I
                        Itm.Focus()
                    Else
                        Dim g As Grid = MyDesigner.Parent
                        MyDesigner.IsEnabled = False
                        MyDesigner.MoveFocus(New TraversalRequest(FocusNavigationDirection.Next))
                        MyDesigner.IsEnabled = True
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
            If Not ApplyRotationMenuItem.IsChecked Then DiagramTextBlock.RenderTransform = DesignerItem.RenderTransform.Inverse
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
        Commands.ChangeBrush(Diagram, Designer.DiagramTextForegroundProperty)
        ApplyLastChangeToSelected()
    End Sub

    Private Sub TextBackgroundMenuItem_Click(sender As Object, e As RoutedEventArgs)
        Commands.ChangeBrush(Diagram, Designer.DiagramTextBackgroundProperty)
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

    Private Sub ApplyRotationMenuItem_Checked(sender As Object, e As RoutedEventArgs)
        If DiagramTextBlock Is Nothing Then Return
        Try
            DiagramTextBlock.RenderTransform = Nothing
            Designer.SetDiagramTextApplyRotation(Diagram, True)
            ApplyLastChangeToSelected()
        Catch
        End Try
    End Sub

    Private Sub ApplyRotationMenuItem_Unchecked(sender As Object, e As RoutedEventArgs)
        DiagramTextBlock.RenderTransform = DesignerItem.RenderTransform.Inverse
        Designer.SetDiagramTextApplyRotation(Diagram, False)
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

    Private Sub CopyMenuItem_Click(sender As Object, e As RoutedEventArgs)
        MyDesigner.Copy()
    End Sub

    Private Sub CutMenuItem_Click(sender As Object, e As RoutedEventArgs)
        MyDesigner.Cut()
    End Sub

    Private Sub DeleteMenuItem_Click(sender As Object, e As RoutedEventArgs)
        MyDesigner.RemoveSelectedItems()
    End Sub

    Private Sub DrawOutlineMenuItem_Checked(sender As Object, e As RoutedEventArgs)
        Dim OldStare As New PropertyState(Diagram, Designer.DiagramTextOutlinedProperty)
        Dim CanUndo As Boolean = True
        If DrawOutlineMenuItem.Tag = "DontUndo" Then
            DrawOutlineMenuItem.Tag = ""
            CanUndo = False
        End If
        Designer.SetDiagramTextOutlined(Diagram, True)
        DiagramObj.OutlineText()
        If CanUndo Then MyDesigner.UndoStack.ReportChanges(New UndoRedoUnit(OldStare.SetNewValue))
    End Sub

    Private Sub DrawOutlineMenuItem_Unchecked(sender As Object, e As RoutedEventArgs)
        Dim OldStare As New PropertyState(Diagram, Designer.DiagramTextOutlinedProperty)
        Dim CanUndo As Boolean = True
        If DrawOutlineMenuItem.Tag = "DontUndo" Then
            DrawOutlineMenuItem.Tag = ""
            CanUndo = False
        End If
        Designer.SetDiagramTextOutlined(Diagram, False)
        Dim Props = Designer.GetDiagramTextFontProps(Diagram)
        DiagramTextBlock.Background = Designer.GetDiagramTextBackground(Diagram)
        DiagramTextBlock.Foreground = Designer.GetDiagramTextForeground(Diagram)
        If CanUndo Then MyDesigner.UndoStack.ReportChanges(New UndoRedoUnit(OldStare.SetNewValue))
    End Sub


    Private Sub DecreaseOutlineThicknessMenuItem_Click(sender As Object, e As RoutedEventArgs)
        Commands.IncreaseTextOutlineThickness(Diagram, -0.1)
        ApplyLastChangeToSelected()
    End Sub

    Private Sub IicreaseOutlineThicknessMenuItem_Click(sender As Object, e As RoutedEventArgs)
        Commands.IncreaseTextOutlineThickness(Diagram, 0.1)
        ApplyLastChangeToSelected()
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
        If BoldMenuItem.IsChecked Then
            Commands.ChangeFontProperty(Diagram, TextBlock.FontWeightProperty, FontWeights.Bold)
        Else
            Commands.ChangeFontProperty(Diagram, TextBlock.FontWeightProperty, FontWeights.Normal)
        End If
        ApplyLastChangeToSelected()
    End Sub

    Private Sub ItalicMenuItem_Checked(sender As Object, e As RoutedEventArgs)
        If ItalicMenuItem.IsChecked Then
            Commands.ChangeFontProperty(Diagram, TextBlock.FontStyleProperty, FontStyles.Italic)
        Else
            Commands.ChangeFontProperty(Diagram, TextBlock.FontStyleProperty, FontStyles.Normal)
        End If
        ApplyLastChangeToSelected()
    End Sub

    Private Sub UnderlineMenuItem_Checked(sender As Object, e As RoutedEventArgs)
        Dim Decorations As TextDecorationCollection = Nothing
        If UnderlineMenuItem.IsChecked Then Decorations = TextDecorations.Underline
        Commands.ChangeFontProperty(Diagram, TextBlock.TextDecorationsProperty, Decorations)
        ApplyLastChangeToSelected()
    End Sub

    Private Sub OutlineFillMenuItem_Click(sender As Object, e As RoutedEventArgs)
        Commands.ChangeBrush(Diagram, Designer.DiagramTextOutlineFillProperty)
        ApplyLastChangeToSelected()
    End Sub


#End Region

    Private Sub ContextMenu_Opened(sender As Object, e As RoutedEventArgs)
        If Me.DiagramGroup Is Nothing Then
            GroupMenuItem.IsEnabled = (MyDesigner.SelectedItems.Count > 1)
            RemoveFromGroupMenuItem.Visibility = Windows.Visibility.Collapsed
        Else
            GroupMenuItem.IsEnabled = True
            If DiagramGroup.Count > 2 Then
                RemoveFromGroupMenuItem.Visibility = Windows.Visibility.Visible
            Else
                RemoveFromGroupMenuItem.Visibility = Windows.Visibility.Collapsed
            End If
        End If
    End Sub

    Friend ExitGroupChecked As Boolean = False

    Private Sub GroupMenuItem_Checked(sender As Object, e As RoutedEventArgs)
        If ExitGroupChecked Then Return

        Dim UndoUnit As New UndoRedoUnit
        Dim GroupTimeStamp = Now
        For Each D As FrameworkElement In MyDesigner.SelectedItems
            Dim A As action = AddressOf DiagramObject.Diagrams(D).AfterRestoreAction
            Dim OldSate As New PropertyState(A, D, Designer.GroupIDProperty)
            Designer.SetGroupID(D, GroupTimeStamp)
            UndoUnit.Add(OldSate.SetNewValue)
        Next
        MyDesigner.UndoStack.ReportChanges(UndoUnit)
        DiagramGroup.Select()
    End Sub

    Private Sub GroupMenuItem_Unchecked(sender As Object, e As RoutedEventArgs)
        If ExitGroupChecked Then Return
        DiagramGroup.Ungroup(Me.DiagramGroup)
    End Sub

    Private Sub RemoveFromGroupMenuItem_Click(sender As Object, e As RoutedEventArgs)
        Designer.SetGroupID(Diagram, Nothing)
    End Sub
End Class
