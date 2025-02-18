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

    Public Property AutoWidth As Boolean
        Get
            Return GetValue(AutoWidthProperty)
        End Get

        Set(ByVal value As Boolean)
            SetValue(AutoWidthProperty, value)
        End Set
    End Property

    Public Shared ReadOnly AutoWidthProperty As DependencyProperty =
                           DependencyProperty.Register("AutoWidth",
                           GetType(Boolean), GetType(DiagramPanel),
                           New PropertyMetadata(False))


    Public Property AutoHeight As Boolean
        Get
            Return GetValue(AutoHeightProperty)
        End Get

        Set(ByVal value As Boolean)
            SetValue(AutoHeightProperty, value)
        End Set
    End Property

    Public Shared ReadOnly AutoHeightProperty As DependencyProperty =
                           DependencyProperty.Register("AutoHeight",
                           GetType(Boolean), GetType(DiagramPanel),
                           New PropertyMetadata(False))

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

        For Each fw As FrameworkElement In Dsn.SelectedItems
            If fw Is Diagram Then Continue For

            For i = 0 To Unit.Count - 1
                Dim propState = TryCast(Unit(i), PropertyState)
                If propState Is Nothing Then Continue For

                Dim target = Commands.GetTarget(fw, propState.Owner)
                If target Is Nothing Then Continue For

                Dim NewPropState As New PropertyState(target, propState.Keys.ToArray)

                For Each pair In propState
                    Dim oldValue = pair.Value.OldValue
                    Dim newValue = pair.Value.NewValue
                    If oldValue Is Nothing OrElse (oldValue IsNot newValue AndAlso Not oldValue.Equals(newValue)) Then
                        target.SetValue(pair.Key, Helper.Clone(pair.Value.NewValue))
                    End If
                Next

                If NewPropState.HasChanges Then Unit.Add(NewPropState.SetNewValues())
            Next
        Next
    End Sub


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
        Dsn.UndoStack.ReportChanges(New UndoRedoUnit(OldState.SetNewValues()))
    End Sub

    Public Sub SendToBack()
        Dsn.MinZIndex -= 1
        Dim OldState = New PropertyState(DesignerItem, Canvas.ZIndexProperty)
        Canvas.SetZIndex(DesignerItem, Dsn.MinZIndex)
        Dsn.UndoStack.ReportChanges(New UndoRedoUnit(OldState.SetNewValues()))
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
            Case Key.F1
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
            UndoUnit.Add(OldSate.SetNewValues)
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


    Public Shared Function GetIsDiagramEnabled(ByVal element As DependencyObject) As Boolean
        If element Is Nothing Then
            Throw New ArgumentNullException("element")
        End If

        Return element.GetValue(IsDiagramEnabledProperty)
    End Function

    Public Shared Sub SetIsDiagramEnabled(ByVal element As DependencyObject, ByVal value As Boolean)
        If element Is Nothing Then
            Throw New ArgumentNullException("element")
        End If

        element.SetValue(IsDiagramEnabledProperty, value)
    End Sub

    Public Shared ReadOnly IsDiagramEnabledProperty As _
               DependencyProperty = DependencyProperty.RegisterAttached("IsDiagramEnabled",
               GetType(Boolean), GetType(FrameworkElement),
               New PropertyMetadata(True))


    Public Shared Function GetIsDiagramVisible(ByVal element As DependencyObject) As Boolean
        If element Is Nothing Then
            Throw New ArgumentNullException("element")
        End If

        Return element.GetValue(IsDiagramVisibleProperty)
    End Function

    Public Shared Sub SetIsDiagramVisible(ByVal element As DependencyObject, ByVal value As Boolean)
        If element Is Nothing Then
            Throw New ArgumentNullException("element")
        End If

        element.SetValue(IsDiagramVisibleProperty, value)
    End Sub

    Public Shared ReadOnly IsDiagramVisibleProperty As _
               DependencyProperty = DependencyProperty.RegisterAttached("IsDiagramVisible",
               GetType(Boolean), GetType(FrameworkElement),
               New PropertyMetadata(True))


    Private Sub PropertiesMenuItem_Click(sender As Object, e As RoutedEventArgs)
        ShowProps()
    End Sub

    Friend Sub ShowProps()
        Dim WndProps As New WndProperties
        With WndProps
            .Show() ' Apply the templates
            .Hide()
            .LeftValue = GetDoubleValue(DesignerItem, Canvas.LeftProperty)
            .TopValue = GetDoubleValue(DesignerItem, Canvas.TopProperty)
            .WidthValue = If(AutoWidth, 0, GetDoubleValue(Me, WidthProperty))
            .HeightValue = If(AutoHeight, 0, GetDoubleValue(Me, HeightProperty))
            .MinWidthValue = GetDoubleValue(Me, MinWidthProperty)
            .MinHeightValue = GetDoubleValue(Me, MinHeightProperty)
            .MaxWidthValue = GetDoubleValue(Me, MaxWidthProperty)
            .MaxHeightValue = GetDoubleValue(Me, MaxHeightProperty)
            .EnabledValue = GetBooleanValue(IsDiagramEnabledProperty)
            .VisibleValue = GetBooleanValue(IsDiagramVisibleProperty)
            .RightToLeftValue = GetFlowDirectionValue()
            .TagValue = GetStringValue(FrameworkElement.TagProperty)
            .ToolTipValue = GetStringValue(FrameworkElement.ToolTipProperty)
        End With

        SetWordWrapValue(WndProps)

        If WndProps.ShowDialog = True Then
            Dim unit As New UndoRedoUnit
            Dim OldState As PropertyState

            Dim leftChanged = WndProps.LeftValue.HasValue
            Dim LeftValue As Double
            If leftChanged Then
                LeftValue = WndProps.LeftValue.Value
            End If

            Dim topChanged = WndProps.TopValue.HasValue
            Dim topValue As Double
            If topChanged Then
                topValue = WndProps.TopValue.Value
            End If

            Dim minWidthChanged = WndProps.MinWidthValue.HasValue
            Dim minWidthValue As Double
            If minWidthChanged Then
                minWidthValue = WndProps.MinWidthValue.Value
            End If

            Dim minHeightChanged = WndProps.MinHeightValue.HasValue
            Dim minHeightValue As Double
            If minHeightChanged Then
                minHeightValue = WndProps.MinHeightValue.Value
            End If

            Dim maxWidthChanged = WndProps.MaxWidthValue.HasValue
            Dim maxWidthValue As Double
            If maxWidthChanged Then
                maxWidthValue = WndProps.MaxWidthValue.Value
            End If

            Dim maxHeightChanged = WndProps.MaxHeightValue.HasValue
            Dim maxHeightValue As Double
            If maxHeightChanged Then
                maxHeightValue = WndProps.MaxHeightValue.Value
            End If

            Dim rtlChanged = WndProps.RightToLeftValue.HasValue
            Dim flowDir As FlowDirection
            If rtlChanged Then
                flowDir = If(WndProps.RightToLeftValue.Value, FlowDirection.RightToLeft, FlowDirection.LeftToRight)
            End If

            Dim toolTipChanged = WndProps.chkToolTip.IsChecked
            Dim toolTipValue As String
            If toolTipChanged Then
                toolTipValue = WndProps.ToolTipValue
            End If

            Dim tagChanged = WndProps.chkTag.IsChecked
            Dim tagValue As String
            If tagChanged Then
                tagValue = WndProps.TagValue
            End If

            For Each fw As FrameworkElement In Dsn.SelectedItems
                If leftChanged OrElse topChanged Then
                    Dim dsnItem = Helper.GetListBoxItem(fw)
                    OldState = New PropertyState(dsnItem)
                    If leftChanged AndAlso Not AreEquals(Canvas.GetLeft(dsnItem), LeftValue) Then
                        OldState.Add(Canvas.LeftProperty)
                        Canvas.SetLeft(dsnItem, LeftValue)
                    End If

                    If topChanged AndAlso Not AreEquals(Canvas.GetTop(dsnItem), topValue) Then
                        OldState.Add(Canvas.TopProperty)
                        Canvas.SetTop(dsnItem, topValue)
                    End If

                    If OldState.HasChanges Then unit.Add(OldState.SetNewValues)
                End If

                If minWidthChanged OrElse minHeightChanged OrElse maxWidthChanged OrElse maxHeightChanged Then
                    Dim pnl = Helper.GetDiagramPanel(fw)
                    OldState = New PropertyState(pnl)

                    If minWidthChanged AndAlso Not AreEquals(pnl.MinWidth, minWidthValue) Then
                        OldState.Add(FrameworkElement.MinWidthProperty)
                        pnl.MinWidth = minWidthValue
                    End If

                    If minHeightChanged AndAlso Not AreEquals(pnl.MinHeight, minHeightValue) Then
                        OldState.Add(FrameworkElement.MinHeightProperty)
                        pnl.MinHeight = minHeightValue
                    End If

                    If maxWidthChanged AndAlso Not AreEquals(pnl.MaxWidth, maxWidthValue) Then
                        OldState.Add(FrameworkElement.MaxWidthProperty)
                        pnl.MaxWidth = maxWidthValue
                    End If

                    If maxHeightChanged AndAlso Not AreEquals(pnl.MaxHeight, maxHeightValue) Then
                        OldState.Add(FrameworkElement.MaxHeightProperty)
                        pnl.MaxHeight = maxHeightValue
                    End If

                    If OldState.HasChanges Then unit.Add(OldState.SetNewValues)
                End If

                OldState = New PropertyState(fw)
                If rtlChanged AndAlso fw.FlowDirection <> flowDir Then
                    OldState.Add(FrameworkElement.FlowDirectionProperty)
                    fw.FlowDirection = flowDir
                End If

                If toolTipChanged AndAlso fw.ToolTip <> toolTipValue Then
                    OldState.Add(FrameworkElement.ToolTipProperty)
                    fw.ToolTip = toolTipValue
                End If

                If tagChanged AndAlso fw.Tag <> tagValue Then
                    OldState.Add(FrameworkElement.TagProperty)
                    fw.Tag = tagValue
                End If

                If OldState.HasChanges Then unit.Add(OldState.SetNewValues)
            Next

            SetDiagramsSize(WndProps, unit)
            SetBooleanValue(DiagramPanel.IsDiagramEnabledProperty, WndProps.EnabledValue, unit)
            SetBooleanValue(DiagramPanel.IsDiagramVisibleProperty, WndProps.VisibleValue, unit)

            If WndProps.WordWrapValue.HasValue Then
                SetWordWrap(WndProps.WordWrapValue, unit)
            End If

            If unit.Count > 0 Then Dsn.UndoStack.ReportChanges(unit)
        End If
    End Sub

    Private Function GetStringValue(dp As DependencyProperty) As String
        Dim value = If(CStr(Diagram.GetValue(dp)), "")
        If Dsn.SelectedItems.Count = 1 Then Return value

        For Each fw As FrameworkElement In Dsn.SelectedItems
            If fw Is Diagram Then Continue For
            Dim v = If(CStr(fw.GetValue(dp)), "")
            If v <> value Then Return Nothing
        Next

        Return value
    End Function

    Private Function GetFlowDirectionValue() As Boolean?
        Dim value = Diagram.FlowDirection
        If Dsn.SelectedItems.Count = 1 Then Return value = FlowDirection.RightToLeft

        For Each fw As FrameworkElement In Dsn.SelectedItems
            If fw Is Diagram Then Continue For
            If fw.FlowDirection <> value Then Return Nothing
        Next

        Return value = FlowDirection.RightToLeft
    End Function

    Private Sub SetBooleanValue(dp As DependencyProperty, boolValue As Boolean?, unit As UndoRedoUnit)
        If boolValue Is Nothing Then Return
        Dim v = boolValue.Value
        For Each fw As FrameworkElement In Dsn.SelectedItems
            Dim OldState As New PropertyState(fw, dp)
            If fw.GetValue(dp) <> v Then fw.SetValue(dp, v)
            If OldState.HasChanges Then unit.Add(OldState.SetNewValues)
        Next
    End Sub

    Private Function GetBooleanValue(dp As DependencyProperty) As Boolean?
        Dim value As Boolean = Diagram.GetValue(dp)
        If Dsn.SelectedItems.Count = 1 Then Return value

        For Each fw As FrameworkElement In Dsn.SelectedItems
            If fw Is Diagram Then Continue For
            If Not fw.GetValue(dp).Equals(value) Then Return Nothing
        Next

        Return value
    End Function

    Private Sub SetDiagramsSize(WndProps As WndProperties, unit As UndoRedoUnit)
        Dim newWidth = WndProps.WidthValue
        Dim newHeight = WndProps.HeightValue
        If newWidth Is Nothing AndAlso newHeight Is Nothing Then Return

        For Each d1 In Dsn.SelectedItems
            SetDiagramSize(Helper.GetDiagramPanel(d1), newWidth, newHeight, unit)
        Next
    End Sub

    Friend Sub SetSize(width As Double, height As Double)
        SetDiagramSize(Me, width, height, New UndoRedoUnit())
    End Sub

    Private Shared Sub SetDiagramSize(pnl As DiagramPanel, newWidth As Double?, newHeight As Double?, unit As UndoRedoUnit)
        Dim diagram = pnl.Diagram
        Dim changeWidth = newWidth.HasValue AndAlso Not AreEquals(pnl.Width, newWidth.Value)
        Dim changeHeight = newHeight.HasValue AndAlso Not AreEquals(pnl.Height, newHeight.Value)

        If Not (changeWidth OrElse changeHeight) Then Return

        Dim propState As New PropertyState(pnl)
        Dim diagram2 As FrameworkElement = Nothing

        If (changeWidth AndAlso Double.IsNaN(newWidth)) OrElse (
                changeHeight AndAlso Double.IsNaN(newHeight)
            ) Then diagram2 = ResizeDiagram(diagram)

        Dim w = -1
        Dim h = -1

        If changeWidth Then
            propState.Add(FrameworkElement.WidthProperty, AutoWidthProperty)
            If Double.IsNaN(newWidth) Then
                pnl.AutoWidth = True
                w = diagram2.ActualWidth
                pnl.Width = w + SystemParameters.ScrollWidth
            Else
                pnl.AutoWidth = False
                pnl.Width = newWidth.Value
            End If
        End If

        If changeHeight Then
            propState.Add(FrameworkElement.HeightProperty, AutoHeightProperty)
            If Double.IsNaN(newHeight) Then
                pnl.AutoHeight = True
                h = diagram2.ActualHeight
                pnl.Height = h + SystemParameters.ScrollHeight
            Else
                pnl.AutoHeight = False
                pnl.Height = newHeight.Value
            End If
        End If

        Helper.UpdateControl(pnl)
        If w > -1 Then pnl.Width = w + 4
        If h > -1 Then pnl.Height = h + 4
        unit.Add(propState.SetNewValues)
    End Sub

    Private Shared Function ResizeDiagram(diagram As FrameworkElement) As FrameworkElement
        Dim d2 = CType(Helper.Clone(diagram), FrameworkElement)
        Dim canv = New Canvas()
        canv.Children.Add(d2)

        Dim tempWnd = New Window() With {
            .ShowInTaskbar = False,
            .WindowStyle = WindowStyle.None,
            .AllowsTransparency = True,
            .Background = Brushes.Transparent,
            .Content = canv
        }

        tempWnd.Show()
        tempWnd.Close()
        tempWnd = Nothing
        Return d2
    End Function

    Private Sub SetWordWrap(wordWrapValue As Boolean, unit As UndoRedoUnit)
        Dim OldState As PropertyState = Nothing
        Dim value = If(wordWrapValue, TextWrapping.Wrap, TextWrapping.NoWrap)

        For Each fw As FrameworkElement In Dsn.SelectedItems
            If TypeOf fw Is TextBox Then
                Dim txt = CType(fw, TextBox)
                If txt.TextWrapping <> value Then
                    OldState = New PropertyState(txt, TextBox.TextWrappingProperty)
                    txt.TextWrapping = value
                    unit.Add(OldState.SetNewValues)
                End If

            ElseIf TypeOf fw Is ContentControl Then
                Dim tb = TryCast(CType(fw, ContentControl).Content, TextBlock)
                If tb IsNot Nothing AndAlso tb.TextWrapping <> value Then
                    OldState = New PropertyState(tb, TextBlock.TextWrappingProperty)
                    tb.TextWrapping = value
                    unit.Add(OldState.SetNewValues)
                End If
            End If
        Next
    End Sub

    Private Sub SetDiagramWordWrap(fw As FrameworkElement, value As Boolean)
        Dim wr = If(value, TextWrapping.Wrap, TextWrapping.NoWrap)

        If TypeOf fw Is TextBox Then
            CType(fw, TextBox).TextWrapping = wr
        ElseIf TypeOf fw Is ContentControl Then
            Dim tb = TryCast(CType(fw, ContentControl).Content, TextBlock)
            If tb IsNot Nothing Then tb.TextWrapping = wr
        End If
    End Sub

    Private Sub SetWordWrapValue(WndProps As WndProperties)
        Dim value As Boolean? = Nothing
        Dim n = 0

        If Dsn.SelectedItems.Count = 1 Then
            value = GetDiagramWordWrap(Diagram)
            If value Is Nothing Then
                WndProps.cmbWordWrap.IsEnabled = False
            Else
                WndProps.WordWrapValue = value
            End If
            Return
        End If

        For Each fw As FrameworkElement In Dsn.SelectedItems
            Dim wr = GetDiagramWordWrap(fw)
            If wr Is Nothing Then
                n += 1
            ElseIf value.HasValue Then
                If wr.Value <> value.Value Then
                    WndProps.WordWrapValue = Nothing
                    Return
                End If
            Else
                value = wr
            End If
        Next

        If n = Dsn.SelectedItems.Count Then
            WndProps.cmbWordWrap.IsEnabled = False
        Else
            WndProps.WordWrapValue = value
        End If

    End Sub

    Private Function GetDiagramWordWrap(control As FrameworkElement) As Boolean?
        If TypeOf control Is TextBox Then
            Return (CType(control, TextBox).TextWrapping = TextWrapping.Wrap)
        ElseIf TypeOf control Is ContentControl Then
            Dim tb = TryCast(CType(control, ContentControl).Content, TextBlock)
            If tb IsNot Nothing Then Return (tb.TextWrapping = TextWrapping.Wrap)
        End If

        Return Nothing
    End Function



    Private Function GetDoubleValue(item As FrameworkElement, dp As DependencyProperty) As Double?
        Dim value As Double = item.GetValue(dp)
        If Dsn.SelectedItems.Count = 1 Then Return value

        Dim onCanvus = TypeOf item Is ListBoxItem
        For Each fw As FrameworkElement In Dsn.SelectedItems
            If fw Is item Then Continue For

            If onCanvus Then
                If Not AreEquals(CDbl(Helper.GetListBoxItem(fw).GetValue(dp)), value) Then Return Nothing
            ElseIf Not AreEquals(CDbl(Helper.GetDiagramPanel(fw).GetValue(dp)), value) Then
                Return Nothing
            End If
        Next

        Return value
    End Function
End Class
