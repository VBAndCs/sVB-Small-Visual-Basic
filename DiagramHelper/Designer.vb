Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.Windows.Markup
Imports System.Xml

Public Class Designer
    Inherits ListBox

    Friend GridLinesBorder As Border
    Friend DesignerCanvas As Canvas
    Friend ConnectionCanvas As Canvas
    Friend TbTopLocation As TextBlock
    Friend TbLeftLocation As TextBlock
    Friend GridBrush As DrawingBrush
    Friend Editor As TextBox
    Friend AllowTransparencyMenuItem As MenuItem
    Friend ZoomBox As ZoomBox
    Friend MaxZIndex As Integer = 0
    Friend Shared Editing As Boolean = False
    Friend SelectionBorder As Border
    Friend WithEvents UndoStack As New UndoRedoStack(Of UndoRedoUnit)(1000)
    Dim DeleteUndoUnit As UndoRedoUnit
    Friend SelectedBounds As Rect
    Friend GridPen As Pen

    Public ReadOnly Property AllowTransparency As Boolean
        Get
            If AllowTransparencyMenuItem Is Nothing Then
                Return OpenFileCanvas IsNot Nothing AndAlso CBool(OpenFileCanvas.Tag)
            End If
            Return AllowTransparencyMenuItem.IsChecked
        End Get
    End Property


    Dim _codeFile As String = ""

    Public Property CodeFile As String
        Get
            Return _codeFile
        End Get

        Set(value As String)
            _codeFile = value
        End Set
    End Property

    Dim _xamlFile As String = ""
    Public Property XamlFile As String
        Get
            Return _xamlFile
        End Get

        Set(value As String)
            _xamlFile = If(value = "", "", IO.Path.GetFullPath(value))
            If _xamlFile <> "" AndAlso CodeFile = "" Then
                _codeFile = _xamlFile.Substring(0, _xamlFile.Length - 5) & ".sb"
            End If
        End Set
    End Property

    Public Sub New()
        Dim resourceLocater As New Uri("/DiagramHelper;component/Resources/designerdecorator.xaml", System.UriKind.Relative)
        Dim ResDec As ResourceDictionary = Application.LoadComponent(resourceLocater)
        Resources.MergedDictionaries.Add(ResDec)
        Style = FindResource("CanvasListBoxStyle")
        ItemContainerStyle = FindResource("listBoxItemStyle")
        SelectionMode = SelectionMode.Multiple
        AllowDrop = True
    End Sub

    Private Shared NewPageOpened As Boolean

    Private Sub Designer_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        If ScrollViewer IsNot Nothing Then Return

        Dim Brdr As Border = VisualTreeHelper.GetChild(Me, 0)
        Dim G = CType(Brdr.Child, Grid)

        ScrollViewer = G.Children(0)
        GridLinesBorder = G.Children(1)
        G = ScrollViewer.Content
        DesignerCanvas = VisualTreeHelper.GetChild(G.Children(0), 0)
        DesignerGrid = G
        ConnectionCanvas = G.Children(1)
        DesignerCanvas.Width = PageWidth
        DesignerCanvas.Height = PageHeight
        ConnectionCanvas.Width = PageWidth
        ConnectionCanvas.Height = PageHeight
        SelectionBorder = ConnectionCanvas.Children(0)

        ShowGridChanged(Me, New DependencyPropertyChangedEventArgs(ShowGridProperty, False, Me.ShowGrid))
        GridBrush = Me.FindResource("GridBrush")
        GridPen = FindResource("GridLinesPen")

        AddHandler ScrollViewer.ScrollChanged, AddressOf ScrollViewer_ScrollChanged

        TbTopLocation = Template.FindName("PART_TopLocation", Me)
        TbLeftLocation = Template.FindName("PART_LeftLocation", Me)
        Editor = Template.FindName("PART_Editor", Me)
        AllowTransparencyMenuItem = Template.FindName("AllowTransparencyMenuItem", Me)


        If Not NewPageOpened Then
            CurrentPage = Me
            Dispatcher.Invoke(Sub() OpenNewPage(False))
            NewPageOpened = True
        End If

        If OpenFileCanvas IsNot Nothing Then
            DesignerCanvas.Background = OpenFileCanvas.Background
            AllowTransparencyMenuItem.IsChecked = CBool(OpenFileCanvas.Tag)
            OpenFileCanvas = Nothing
        End If

        Focus()

    End Sub

    Friend SelectedToolBoxItem As ToolBoxItem

    Friend Property LocationVisibility As Windows.Visibility
        Get
            Return TbLeftLocation.Visibility
        End Get

        Set(value As Windows.Visibility)
            TbLeftLocation.Visibility = value
            TbTopLocation.Visibility = value
        End Set
    End Property

    Public Function SetControlName(controlIndex As Integer, name As String) As Boolean
        If controlIndex = -1 Then
            Return ChangeFormName(name)
        Else
            Return SetControlName(Me.Items(controlIndex), name)
        End If
    End Function

    Public Function SetControlName(control As UIElement, name As String) As Boolean
        If name = "" Then Return False

        Dim newName = name.ToLower()
        If Me.GetControlName(control).ToLower() <> newName Then
            If IsNumeric(newName(0)) Then
                MessageBox.Show("Name can't start with a number.", "Invalid Name", MessageBoxButton.OK, MessageBoxImage.Error)
                Return False
            End If

            If newName = Me.Name.ToLower() Then
                MessageBox.Show($"'{newName}' is the form name! Choose another name", "Invalid Name", MessageBoxButton.OK, MessageBoxImage.Error)
                Return False
            End If

            For Each cnt In Me.Items
                If Me.GetControlName(cnt).ToLower() = newName Then
                    MessageBox.Show($"There is another control named '{newName}'!", "Invalid Name", MessageBoxButton.OK, MessageBoxImage.Error)
                    Return False
                End If
            Next
        End If

        Try
            Dim OldState As New PropertyState(
                Sub()
                    Me.SelectedItem = control
                    SetControlName(control, Automation.AutomationProperties.GetName(control))
                    RaiseEvent PageShown(UpdateControlNameAndText)
                End Sub,
                control,
                Automation.AutomationProperties.NameProperty)

            Dim fw = TryCast(control, FrameworkElement)
            If fw IsNot Nothing Then fw.Name = name
            Automation.AutomationProperties.SetName(control, name)

            Me.UndoStack.ReportChanges(New UndoRedoUnit(OldState.SetNewValue))
            RaiseEvent PageShown(UpdateControlNameAndText)
            Return True

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Invalid Name", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try

        Return False
    End Function

    Public Function ChangeFormName(name As String) As Boolean
        If name = "" Then Return False

        Dim newName = name.ToLower()
        If Me.Name.ToLower() <> newName Then
            If IsNumeric(newName(0)) Then
                MessageBox.Show("Name can't start with a number.", "Invalid Name", MessageBoxButton.OK, MessageBoxImage.Error)
                Return False
            End If

            For Each cnt In Me.Items
                If Me.GetControlName(cnt).ToLower() = newName Then
                    MessageBox.Show($"There is a control named '{newName}'!", "Invalid Name", MessageBoxButton.OK, MessageBoxImage.Error)
                    Return False
                End If
            Next
        End If


        Try
            Dim OldState As New PropertyState(AddressOf UpdateFormInfo, Me, Designer.NameProperty)
            Me.Name = name
            Me.UndoStack.ReportChanges(New UndoRedoUnit(OldState.SetNewValue))
            UpdateFormInfo()
            RaiseEvent PageShown(UpdateFormName)
            Return True

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Invalid Name", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try

        Return False

    End Function

    Public Event DiagramDoubleClick(diagram As UIElement)

    Friend Sub OnDiagramDoubleClick(diagram As UIElement)
        RaiseEvent DiagramDoubleClick(diagram)
    End Sub


#Region "Pages"
    Public Shared Pages As New Dictionary(Of String, Designer)
    Public Shared FormNames As New ObservableCollection(Of String)
    Public Shared FormKeys As New List(Of String)
    Public Shared CurrentPage As Designer
    Public Shared Property PagesGrid As Grid
    Public Shared TempProjectPath As String

    Private Shared TempKeyNum As Integer = 0
    Public PageKey As String

    Public Shared Event PageShown(index As Integer)

    Public ReadOnly Property IsNew As Boolean
        Get
            Return Not HasChanges AndAlso _xamlFile = ""
        End Get
    End Property

    Public Shared Function GetTempFormName() As String
        TempKeyNum += 1
        Return "Form" & TempKeyNum
    End Function

    Private Shared globalKey As String
    Private Shared appDir = System.AppDomain.CurrentDomain.BaseDirectory

    Private Shared Function GetTempKey(Optional pagePath As String = "") As String
        If pagePath <> "" Then
            pagePath = pagePath.ToLower()
            If IO.Path.GetFileNameWithoutExtension(pagePath) = "global" Then
                If globalKey <> "" Then Return globalKey
                TempKeyNum += 1
                globalKey = "KEY" & TempKeyNum
                CurrentPage.IsEnabled = False
                Return globalKey
            Else
                For Each p In Pages
                    If p.Value._xamlFile.ToLower() = pagePath Then
                        Return p.Key
                    End If
                Next
            End If
        End If

        TempKeyNum += 1
        Return "KEY" & TempKeyNum
    End Function

    Private Shared Sub UpdatePageInfo(Optional updateForms As Boolean = True)
        If CurrentPage.PageKey = "" Then CurrentPage.PageKey = GetTempKey()
        Pages(CurrentPage.PageKey) = CurrentPage
        If updateForms Then UpdateFormInfo()
    End Sub

    Private Shared Sub UpdateFormInfo()
        Dim displayName = " ● " & CurrentPage.Name & If(CurrentPage._xamlFile = "", " *", "")
        Dim i = FormKeys.IndexOf(CurrentPage.PageKey)
        If i = -1 Then
            If CurrentPage.Name <> "global" Then
                FormNames.Add(displayName)
                FormKeys.Add(CurrentPage.PageKey)
            End If
        ElseIf FormNames(i) <> displayName Then
            FormNames(i) = displayName
        End If
    End Sub

    Private Shared Sub CreateNewDesigner()
        CurrentPage = New Designer()
        PagesGrid.Children.Add(CurrentPage)
        Helper.UpdateControl(CurrentPage)

    End Sub

    Public Shared Function OpenNewPage(Optional UpdateCurrentPage As Boolean = True) As String
        If UpdateCurrentPage Then UpdatePageInfo(False)

        If NewPageOpened Then CreateNewDesigner()
        CurrentPage.PageKey = GetTempKey()
        CurrentPage.Name = CurrentPage.PageKey.Replace("KEY", "Form")
        UpdatePageInfo()

        Call SetDefaultPropertiesSub()
        BringPageToFront()
        CurrentPage.Focus()
        RaiseEvent PageShown(FormKeys.Count - 1)

        Return CurrentPage.PageKey
    End Function

    Public Delegate Sub SetDefaultPropertiesHandler()
    Public Shared SetDefaultPropertiesSub As SetDefaultPropertiesHandler = AddressOf SetDefaultProperties

    Public Shared Sub SetDefaultProperties()
        CurrentPage._Text = CurrentPage.Name
        CurrentPage.DesignerCanvas.Background = SystemColors.ControlBrush
        CurrentPage.GridPen.Thickness = 0.3
        CurrentPage.ShowGrid = True
        CurrentPage.GridPen.Brush = Brushes.LightGray
        CurrentPage.PageWidth = 700
        CurrentPage.PageHeight = 500
    End Sub

    Public Shared Function ClosePage(
                    Optional openNewPageIfLast As Boolean = True,
                    Optional force As Boolean = False,
                    Optional ClosingApp As Boolean = False
               ) As Boolean

        If CurrentPage Is Nothing Then Return True

        If Not force AndAlso CurrentPage.IsNew AndAlso FormKeys.Count = 1 Then
            Return True
        End If

        If CurrentPage.IsDirty Then
            If Not CurrentPage.AskToSave() Then Return False
        End If

        If globalKey <> "" AndAlso CurrentPage.PageKey = globalKey Then
            If ClosingApp Then
                Pages.Remove(globalKey)
                globalKey = ""
            End If

            Dim i = FormKeys.Count - 1
            If i < 0 Then
                If openNewPageIfLast Then
                    OpenNewPage(False)
                Else
                    CurrentPage = Nothing
                    FormNames.Clear()
                End If
            Else
                SwitchTo(FormKeys(i), False)
            End If

        Else
            Pages.Remove(CurrentPage.PageKey)
            PagesGrid.Children.Remove(CurrentPage)

            Dim i = FormKeys.IndexOf(CurrentPage.PageKey)

            If i <> -1 Then
                FormKeys.RemoveAt(i)
                FormNames.RemoveAt(i)
            End If

            Dim nextKey As String
            Dim n = FormKeys.Count - 1
            If n > -1 Then
                If i > n Then
                    nextKey = FormKeys(n)
                Else
                    nextKey = FormKeys(i)
                End If

                SwitchTo(nextKey, False)

            ElseIf openNewPageIfLast Then
                If ProjectExplorer.CurrentProject <> "" Then
                    SwitchTo(Helper.GlobalFileName, False)
                Else
                    OpenNewPage(False)
                End If
            Else
                CurrentPage = Nothing
                FormNames.Clear()
            End If
        End If
        Return True
    End Function

    Public Shared Function SwitchTo(
                    key As String,
                    Optional UpdateCurrentPage As Boolean = True
                ) As String

        If key = "" Then Return OpenNewPage()

        If CurrentPage IsNot Nothing Then
            If CurrentPage.PageKey = key Then Return key
            If CurrentPage._xamlFile.ToLower() = key.ToLower() Then Return key
            If UpdateCurrentPage Then UpdatePageInfo()
        End If

        If globalKey <> "" AndAlso key.ToLower() = Helper.GlobalFileName Then
            CurrentPage = Pages(globalKey)

        ElseIf Pages.ContainsKey(key) Then
            CurrentPage = Pages(key)

        ElseIf Not IO.File.Exists(key) AndAlso key.ToLower() <> Helper.GlobalFileName Then ' closed
            Return key
        Else
            Dim xamlPath = key
            Dim codePath = xamlPath.Substring(0, xamlPath.Length - 5) & ".sb"

            For Each item In Pages
                Dim page = item.Value
                If page._xamlFile.ToLower() = xamlPath OrElse
                             page._codeFile.ToLower() = codePath Then
                    ' File is already opened. Stwich to it.
                    Return SwitchTo(item.Key, False)
                End If
            Next

            ' File is not opened. Open it.
            Open(key)
        End If

        BringPageToFront()
        RaiseEvent PageShown(
            If(key = Helper.GlobalFileName OrElse key = globalKey,
                GlobalFileIndex,
                FormKeys.IndexOf(CurrentPage.PageKey)
            )
        )
        Return CurrentPage.PageKey
    End Function

    Private Shared Sub BringPageToFront()
        Dim z = 0
        For Each item In PagesGrid.Children
            z = Math.Max(z, Panel.GetZIndex(item))
        Next

        Panel.SetZIndex(CurrentPage, z + 1)
    End Sub

    Private Function PageToXaml() As String
        For Each Diagram In Me.Items
            Commands.UpdateFontProperties(Diagram)
        Next

        Dim canvas As New Canvas With {
            .Name = If(Name = "", PageKey.Replace("KEY", "Form"), Name),
            .Width = PageWidth,
            .Height = PageHeight,
            .Background = DesignerCanvas.Background,
            .Tag = AllowTransparencyMenuItem.IsChecked
        }

        If _Text <> "" Then
            Automation.AutomationProperties.SetHelpText(canvas, _Text)
        Else
            canvas.ClearValue(Automation.AutomationProperties.HelpTextProperty)
        End If

        For Each diagram In Me.Items
            Dim diagram2 As FrameworkElement = Helper.Clone(diagram)
            ' Note: These properties are changed by code. 
            ' Keep them if the user changed them in properties window.
            diagram2.ClearValue(AllowDropProperty)
            diagram2.ClearValue(CursorProperty)
            diagram2.ClearValue(IsTabStopProperty)

            ' GroupID, DiagramText will cause isuues in Small Basic, because they need a reference to DiagramHelper
            ' Use any wpf built-in properety to hold theor values

            Dim gID = GetGroupID(diagram2)
            If gID > 0 Then
                Automation.AutomationProperties.SetAutomationId(diagram2, gID)
                diagram2.ClearValue(GroupIDProperty)
            Else
                diagram2.ClearValue(Automation.AutomationProperties.AutomationIdProperty)
            End If


            ' Copy properties from designer to Canvas
            Canvas.SetLeft(diagram2, Designer.GetLeft(diagram2))
            diagram2.ClearValue(Designer.LeftProperty)

            Canvas.SetTop(diagram2, Designer.GetTop(diagram2))
            diagram2.ClearValue(Designer.TopProperty)


            diagram2.Width = Designer.GetFrameWidth(diagram2)
            diagram2.Height = Designer.GetFrameHeight(diagram2)

            diagram2.ClearValue(Designer.FrameWidthProperty)
            diagram2.ClearValue(Designer.FrameHeightProperty)
            diagram2.ClearValue(Designer.DiagramTextFontPropsProperty)
            Dim content = TryCast(diagram2, ContentControl)?.Content
            TryCast(content, TextBlock)?.ClearValue(Designer.DiagramTextFontPropsProperty)

            Dim z = Panel.GetZIndex(Helper.GetListBoxItem(diagram))
            Panel.SetZIndex(diagram2, z)

            Dim angle = Designer.GetRotationAngle(diagram2)
            If angle <> 0 Then Helper.Rotate(diagram2, angle)
            diagram2.ClearValue(RotationAngleProperty)

            Dim skew = TryCast(diagram2.LayoutTransform, SkewTransform)
            If skew IsNot Nothing Then
                If Math.Round(skew.AngleX, 4) = 0 AndAlso Math.Round(skew.AngleY, 4) = 0 Then
                    diagram2.ClearValue(LayoutTransformProperty)
                End If
            End If
            canvas.Children.Add(diagram2)
        Next

        Return XamlWriter.Save(canvas)

    End Function


    Dim OpenFileCanvas As Canvas
    Private Sub XamlToPage(xaml As String)
        If xaml = "" Then Return

        Dim obj = XamlReader.Load(XmlReader.Create(New IO.StringReader(xaml)))
        Dim canvas = TryCast(obj, Canvas)

        If canvas Is Nothing Then
            canvas = New Canvas
            canvas.Name = "UnKnownForm"
            Try
                canvas.Children.Add(obj)
            Catch ex As Exception
                canvas.Children.Add(New Label() With {.Content = "An error hapeened while loading the file."})
            End Try
        End If

        Me.Name = canvas.Name
        _Text = Automation.AutomationProperties.GetHelpText(canvas)

        Me.Visibility = Visibility.Hidden
        If Not Double.IsNaN(canvas.Width) Then Me.PageWidth = canvas.Width
        If Not Double.IsNaN(canvas.Height) Then Me.PageHeight = canvas.Height

        If DesignerCanvas Is Nothing Then
            OpenFileCanvas = canvas
        Else
            DesignerCanvas.Background = canvas.Background
            AllowTransparencyMenuItem.IsChecked = CBool(canvas.Tag)
        End If

        For Each child In canvas.Children
            Dim Diagram = TryCast(Helper.Clone(child), FrameworkElement)
            If Diagram Is Nothing Then Continue For

            Me.Items.Add(Diagram)
            ' Restore the GroupID and DiagramText properties
            Dim txt = Automation.AutomationProperties.GetHelpText(Diagram)
            If txt <> "" Then
                SetControlText(Diagram, txt)
                Diagram.ClearValue(Automation.AutomationProperties.HelpTextProperty)
            End If

            Dim gId = Automation.AutomationProperties.GetAutomationId(Diagram)
            If gId = "" Then
                Diagram.ClearValue(GroupIDProperty)
            Else
                SetGroupID(Diagram, gId)
                Diagram.ClearValue(Automation.AutomationProperties.AutomationIdProperty)
            End If

            Designer.SetFrameWidth(Diagram, Diagram.Width)
            Diagram.ClearValue(WidthProperty)

            Designer.SetFrameHeight(Diagram, Diagram.Height)
            Diagram.ClearValue(HeightProperty)

            Dim left = Canvas.GetLeft(Diagram)
            If Double.IsNaN(left) Then left = 0
            SetLeft(Diagram, left)

            Dim top = Canvas.GetTop(Diagram)
            If Double.IsNaN(top) Then top = 0
            Designer.SetTop(Diagram, top)

            Dim RotateTransform = TryCast(Diagram.RenderTransform, RotateTransform)
            If RotateTransform Is Nothing Then
                Designer.SetRotationAngle(Diagram, 0)
            Else
                Dim angle = RotateTransform.Angle
                Diagram.RenderTransform = Nothing
                Designer.SetRotationAngle(Diagram, angle)
            End If

            Dim lt = Diagram.LayoutTransform
            Diagram.LayoutTransform = Nothing
            Helper.UpdateControl(Diagram)
            Diagram.LayoutTransform = lt
        Next

        Helper.UpdateControl(Me)

        Me.Visibility = Visibility.Visible
    End Sub

#End Region


    Function GetSelectionBounds() As Rect
        Dim MinX As Double = Double.MaxValue
        Dim MaxX As Double = Double.MinValue
        Dim MinY As Double = Double.MaxValue
        Dim MaxY As Double = Double.MinValue

        For Each Diagram As FrameworkElement In Me.SelectedItems
            Dim R As New Rect(0, 0, Diagram.ActualWidth, Diagram.ActualHeight)
            R = Diagram.TransformToVisual(Me.DesignerCanvas).TransformBounds(R)

            If R.TopLeft.X < MinX Then MinX = R.TopLeft.X
            If R.TopRight.X > MaxX Then MaxX = R.TopRight.X

            If R.TopLeft.Y < MinY Then MinY = R.TopLeft.Y
            If R.BottomLeft.Y > MaxY Then MaxY = R.BottomLeft.Y
        Next
        Return New Rect(New Point(MinX, MinY), New Point(MaxX, MaxY))
    End Function

    Private Sub Designer_Drop(sender As Object, e As DragEventArgs) Handles Me.Drop
        Dim Pos = e.GetPosition(Me.DesignerCanvas)
        Dim tbItem As ToolBoxItem = e.Data.GetData(GetType(ToolBoxItem))
        DarwDiagram(tbItem, Pos)
    End Sub

    Private Sub DarwDiagram(
                        tbItem As ToolBoxItem,
                        pos As Point,
                        Optional width As Double = -1,
                        Optional height As Double = -1
                )

        If tbItem IsNot Nothing Then
            Dim newItem = tbItem.Child
            Dim diagram As UIElement
            Dim defaultName = ""
            Dim controlType As Type
            Dim typeName As String
            Dim sbControl = TryCast(newItem, SBControl)

            If sbControl Is Nothing Then
                diagram = Helper.Clone(newItem)
                controlType = newItem.GetType()
                typeName = tbItem.Name
            Else
                diagram = Helper.Clone(sbControl.Control)
                controlType = diagram.GetType()
                typeName = controlType.Name
            End If

            If width > -1 Then SetFrameWidth(diagram, Math.Max(20, width))
            If height > -1 Then SetFrameHeight(diagram, Math.Max(20, height))

            defaultName = GetDefaultControlName(controlType, typeName)

            If defaultName <> "" Then
                Automation.AutomationProperties.SetName(diagram, defaultName)
                SetControlText(diagram, defaultName, False)
            End If

            diagram.ClearValue(ToolTipProperty)


            Dim OldState = New CollectionState(AddressOf AfterRestoreAction, Me.Items, diagram)
            Me.Items.Add(diagram)
            UndoStack.ReportChanges(New UndoRedoUnit(OldState.SetNewValue))

            Designer.SetLeft(diagram, pos.X)
            Designer.SetTop(diagram, pos.Y)
            Helper.UpdateControl(Me)

            Dim Item = Helper.GetListBoxItem(diagram)
            Me.SelectedIndex = -1
            If Item IsNot Nothing Then
                Item.IsSelected = True
                Item.Focus()
            End If
        End If
    End Sub

    Private Function GetDefaultControlName(controlType As Type, typeName As String) As String
        Dim num = 1
        For Each dg In Me.Items
            If dg.GetType() Is controlType Then
                Dim controlName = GetControlName(dg)
                If controlName.StartsWith(typeName) Then
                    Dim n = controlName.Substring(typeName.Length)
                    If IsNumeric(n) Then
                        If CInt(n) >= num Then num = CInt(n) + 1
                    End If
                End If
            End If
        Next
        Return typeName & num
    End Function

    Sub RemoveDiagram(Diagram As UIElement)
        Dim Pnl = Helper.GetDiagramPanel(Diagram)
        DiagramGroup.RemovePanelOnly(Pnl)
        Me.Items.Remove(Diagram)
    End Sub

    Sub RemoveSelectedDiagrams()
        Dim Diagrams = Me.SelectedItems
        Dim OldState As New CollectionState(AddressOf AfterRestoreAction, Me.Items)
        For i = Diagrams.Count - 1 To 0 Step -1
            OldState.Add(Diagrams(i))
            Me.RemoveDiagram(Diagrams(i))
        Next

        If DeleteUndoUnit Is Nothing Then DeleteUndoUnit = New UndoRedoUnit
        DeleteUndoUnit.Add(OldState.SetNewValue)

    End Sub

    Public Sub RemoveSelectedItems()
        DeleteUndoUnit = New UndoRedoUnit
        Me.RemoveSelectedDiagrams()
        Me.UndoStack.ReportChanges(DeleteUndoUnit)
        DeleteUndoUnit = Nothing
        Me.Focus()
    End Sub

    Protected Overrides Sub OnItemsChanged(e As Specialized.NotifyCollectionChangedEventArgs)
        MyBase.OnItemsChanged(e)
        If e.Action = Specialized.NotifyCollectionChangedAction.Remove Then
            For Each Diagram In e.OldItems
                Dim Pnl = Helper.GetDiagramPanel(Diagram)
                Pnl.Dispose()
            Next
        End If
    End Sub

    Public Sub Copy()
        If Me.SelectedIndex = -1 Then Return
        Dim Lst As New ArrayList
        For Each Diagram In Me.SelectedItems
            Commands.UpdateFontProperties(Diagram)
            Lst.Add(Diagram)
        Next
        Dim xaml = XamlWriter.Save(Lst)
        Clipboard.SetData(DataFormats.Xaml, xaml)
    End Sub

    Public Sub Cut()
        DeleteUndoUnit = New UndoRedoUnit
        Me.Copy()
        Me.RemoveSelectedDiagrams()
        Me.UndoStack.ReportChanges(DeleteUndoUnit)
        DeleteUndoUnit = Nothing
        Me.Focus()
    End Sub

    Public Sub Paste()
        Try
            Dim xaml As String = Clipboard.GetData(DataFormats.Xaml)
            Dim Lst As ArrayList = XamlReader.Load(XmlReader.Create(New IO.StringReader(xaml)))
            Me.SelectedItems.Clear()
            Dim OldState = New CollectionState(AddressOf AfterRestoreAction, Me.Items)
            For Each Diagram As UIElement In Lst
                Designer.SetLeft(Diagram, Designer.GetLeft(Diagram) + 10)
                Designer.SetTop(Diagram, Designer.GetTop(Diagram) + 10)
                Dim name = GetControlName(Diagram)
                SetControlName(Diagram, GetNextName(name))
                OldState.Add(Diagram)
                Me.Items.Add(Diagram)
                OldState.SetNewValue()
            Next

            UndoStack.ReportChanges(New UndoRedoUnit(OldState))
            Helper.UpdateControl(Me)

            For Each Diagram As UIElement In Lst
                Dim Item = Helper.GetListBoxItem(Diagram)
                Item?.Focus()
                Me.ScrollIntoView(Diagram)
            Next

            xaml = XamlWriter.Save(Lst)
            Clipboard.SetData(DataFormats.Xaml, xaml)

        Catch ex As Exception

        End Try
    End Sub

    Private Function GetNextName(name As String) As String
        Dim baseName As String
        Dim num As Integer
        For i = name.Length - 1 To 0 Step -1
            If Not IsNumeric(name(i)) Then
                If i = name.Length - 1 Then
                    baseName = name
                    num = 0
                Else
                    baseName = name.Substring(0, i + 1)
                    num = CInt(name.Substring(i + 1))
                End If
                Exit For
            End If
        Next

        For Each dg In Me.Items
            Dim controlName = GetControlName(dg)
            If controlName.StartsWith(baseName) Then
                Dim n = controlName.Substring(baseName.Length)
                If n = "" Then
                    If num = 0 Then num = 1
                ElseIf IsNumeric(n) Then
                    If CInt(n) >= num Then num = CInt(n) + 1
                End If
            End If
        Next
        Return baseName & If(num = 0, "", CStr(num))

    End Function

    Public Function CanPaste() As Boolean
        If Not Clipboard.ContainsData(DataFormats.Xaml) Then Return False
        Dim xaml As String = Clipboard.GetData(DataFormats.Xaml)
        Dim Lst = TryCast(XamlReader.Load(XmlReader.Create(New IO.StringReader(xaml))), ArrayList)
        If Lst Is Nothing Then Return False
        Return True
    End Function


    Public Function Save() As Boolean
        If _xamlFile = "" Then
            Return SaveAs()
        ElseIf Me.HasChanges Then
            Return SavePage(_xamlFile, False)
        Else
            Return True
        End If
    End Function

    Public Shared RecentDirectory As String

    Public Function SaveAs() As Boolean
        ' Configure open file dialog box 
        Dim dlg As New Microsoft.Win32.SaveFileDialog With {
            .DefaultExt = ".xaml", ' Default file extension
            .Filter = "Diagram Pages|*.xaml",
            .Title = "Save Diagram Design Page"
        }

        Dim saveName As String
        If _xamlFile = "" Then
            dlg.InitialDirectory = RecentDirectory
            saveName = Me.Name
        Else
            dlg.InitialDirectory = _xamlFile
            saveName = IO.Path.GetFileNameWithoutExtension(_xamlFile)
        End If
        dlg.FileName = saveName

        If dlg.ShowDialog() = True Then
            Dim oldPath = _xamlFile
            _xamlFile = dlg.FileName ' don't use FileName property to keep the old code file path
            If Not SavePage(oldPath, True) Then Return False
            XamlFile = _xamlFile ' Update the old code file path
            Return True
        End If
        Return False
    End Function

    Public Shared Function SavePageIfDirty(xamlFile As String) As Boolean
        Dim formFile = xamlFile.ToLower()

        For Each page In Pages
            Dim dsn = page.Value
            If dsn._xamlFile.ToLower() = formFile Then
                If dsn.HasChanges Then
                    SwitchTo(page.Key)
                    CurrentPage.Save()
                    Return True
                End If
            End If
        Next

        Return False
    End Function

    Public Delegate Function SavePageDelegate(oldPath As String, saveAs As Boolean) As Boolean
    Public SavePage As SavePageDelegate = AddressOf DoSave

    Public Function DoSave(Optional tmpPath As String = Nothing, Optional saveAs As Boolean = False) As Boolean
        Try
            Dim xmal = PageToXaml()
            Dim saveTo = If(tmpPath = "", _xamlFile, tmpPath)
            IO.File.WriteAllText(saveTo, xmal, System.Text.Encoding.UTF8)
            _codeFile = saveTo.Substring(0, saveTo.Length - 5) & ".sb"
            If tmpPath = "" Then
                UpdateFormInfo()
                Me.HasChanges = False
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
            Return False
        End Try
        Return True
    End Function

    Public Sub SaveToImage()
        Dim Sc = Me.Scale
        Me.Scale = 1
        Me.ScrollViewer.ScrollToHorizontalOffset(0)
        Me.ScrollViewer.ScrollToVerticalOffset(0)
        Me.SelectedIndex = -1
        Me.Focus()
        Dim ImgSaver As New ImageSaver
        ImgSaver.Save(DesignerGrid, _xamlFile)
        Me.Scale = Sc
    End Sub

    Public Sub Open()
        ' Configure open file dialog box 
        Dim dlg As New Microsoft.Win32.OpenFileDialog With {
            .DefaultExt = ".xaml", ' Default file extension
            .Filter = "Diagram Pages|*.xaml",
            .Title = "Open Diagram Design Page"
        }
        If dlg.ShowDialog() = True Then
            If Not CurrentPage.HasChanges AndAlso
                    CurrentPage.IsNew AndAlso Pages.Count = 1 Then
                ClosePage(False, True)
            End If

            SwitchTo(dlg.FileName)
        End If
    End Sub


    Public Shared Sub Open(fileName As String)
        If fileName.ToLower() = Helper.GlobalFileName Then
            fileName = IO.Path.Combine(appDir, Helper.ExactGlobalFileName)
        End If

        If Not IO.File.Exists(fileName) Then
            MsgBox($"File '{fileName}' doesn't exist!")
            Return
        End If

        PagesGrid.Cursor = Cursors.Wait
        Try
            Dim xaml = IO.File.ReadAllText(fileName, System.Text.Encoding.UTF8)
            ' Ensure we don't open new page after opneing a file from command line args
            If NewPageOpened Then
                CreateNewDesigner()
            Else
                CurrentPage = PagesGrid.Children(0)
                NewPageOpened = True
            End If

            CurrentPage.XamlToPage(xaml)
            CurrentPage.ShowGrid = True
            CurrentPage._xamlFile = IO.Path.GetFullPath(fileName)
            CurrentPage.PageKey = GetTempKey(CurrentPage._xamlFile)
            Pages(CurrentPage.PageKey) = CurrentPage
            UpdateFormInfo()

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            PagesGrid.Cursor = Nothing
        End Try
    End Sub


    Private Function AskToSave() As Boolean
        Select Case MessageBox.Show($"'{Me.Name}' has changed. Do you want to save changes?", "Save Changes", MessageBoxButton.YesNoCancel)
            Case MessageBoxResult.Yes
                Return Me.Save()
            Case MessageBoxResult.No
                Return True
            Case MessageBoxResult.Cancel
                Return False
        End Select

        Return True
    End Function


    Public Sub Print()
        Dim Sc = Me.Scale
        Me.Scale = 1
        Me.ScrollViewer.ScrollToHorizontalOffset(0)
        Me.ScrollViewer.ScrollToVerticalOffset(0)
        Me.SelectedIndex = -1
        Me.Focus()
        Dim dialog As New PrintDialog()
        If dialog.ShowDialog() = True Then
            dialog.PrintVisual(Me.DesignerGrid, _xamlFile)
        End If
        Me.Scale = Sc
    End Sub

    Public Sub Undo()
        If UndoStack.CanUndo Then UndoStack.Undo().RestoreOldState()
    End Sub

    Public Sub Redo()
        If UndoStack.CanRedo Then UndoStack.Redo().RestoreNewState()
    End Sub

    Public Shadows Sub SelectAll()
        MyBase.SelectAll()
        Dim Item = Helper.GetListBoxItem(Me.SelectedItem)
        If Item IsNot Nothing Then Item.Focus()
    End Sub

    Public Sub IncreaseGridThickness(Value As Single)
        If (Value < 0 AndAlso GridPen.Thickness <= 0.1) OrElse (Value > 0 AndAlso GridPen.Thickness >= 1.5) Then Return
        Dim OldState As New PropertyState(GridPen, Pen.ThicknessProperty)
        GridPen.Thickness += Value
        Me.UndoStack.ReportChanges(New UndoRedoUnit(OldState.SetNewValue))
    End Sub

    Friend Sub AfterRestoreAction(Diagrams As IList)
        Helper.UpdateControl(Me)
        Me.SelectedIndex = -1
        For Each Diagram In Diagrams
            Dim Item = Helper.GetListBoxItem(Diagram)
            If Item Is Nothing Then ' Deleted
                If Me.Items.Count > 0 Then
                    Dim Last = Me.Items(Me.Items.Count - 1)
                    Dim LastItem = Helper.GetListBoxItem(Last)
                    LastItem.Focus()
                Else
                    Me.Focus()
                End If
            Else
                'Item.IsSelected = True
                Item.Focus()
            End If
        Next
    End Sub

    Private Sub Designer_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles Me.MouseDown
        If Not Editing AndAlso e.Source Is Me AndAlso
                    Keyboard.Modifiers <> ModifierKeys.Control Then
            Me.SelectedIndex = -1
            Me.Focus()
        End If
    End Sub

    Dim FirstTime As Boolean = True
    Private Sub ScrollViewer_ScrollChanged(sender As Object, e As ScrollChangedEventArgs)
        If FirstTime Then
            FirstTime = False
            Me.UpdateGrid()
        End If

        GridBrush.Viewport = New Rect(
               -(e.HorizontalOffset / Me.Scale) Mod Helper.CmToPx,
               -(e.VerticalOffset / Me.Scale) Mod Helper.CmToPx,
               Helper.CmToPx, Helper.CmToPx
        )

    End Sub

    Private Sub Designer_PreviewKeyDown(sender As Object, e As KeyEventArgs) Handles Me.PreviewKeyDown
        If e.OriginalSource Is Editor Then Return

        If e.Key = Input.Key.Escape Then
            If Me.SelectedToolBoxItem IsNot Nothing Then
                SelectedToolBoxItem.IsSelected = False
            Else
                Me.SelectedIndex = -1
            End If

            Return
        End If

        If Keyboard.Modifiers = ModifierKeys.Control Then
            Select Case e.Key
                Case Key.OemPlus
                    IncreaseGridThickness(0.1)
                Case Key.OemMinus
                    IncreaseGridThickness(-0.1)
                Case Key.L
                    Commands.ChangeBrush(GridPen, Pen.BrushProperty)
                Case Key.Z
                    Me.Undo()
                    e.Handled = True
                    Return
                Case Key.Y
                    Me.Redo()
                    e.Handled = True
                    Return
                Case Key.A
                    Me.SelectAll()
                    e.Handled = True
                    Return
                Case Key.C
                    Me.Copy()
                    e.Handled = True
                    Return
                Case Key.X
                    Me.Cut()
                    e.Handled = True
                    Return
                Case Key.V
                    Me.Paste()
                    e.Handled = True
                    Return
                Case Key.S
                    Me.Save()
                    e.Handled = True
                    Return
                Case Key.M
                    Me.SaveToImage()
                    e.Handled = True
                    Return
                Case Key.O
                    Me.Open()
                    e.Handled = True
                    Return
                Case Key.N
                    Me.OpenNewPage()
                    e.Handled = True
                    Return
                Case Key.P
                    Me.Print()
                    e.Handled = True
                    Return
                Case Key.F4
                    ClosePage()
                    e.Handled = True
                    Return
            End Select
        End If

        If e.Key = Key.Delete Then
            Me.RemoveSelectedItems()
            e.Handled = True
            Return
        End If

        If Me.SelectedIndex > -1 Then Return

        Select Case e.Key
            Case Key.Up
                If Keyboard.Modifiers = ModifierKeys.Control Then
                    ScrollViewer.ScrollToVerticalOffset(ScrollViewer.VerticalOffset - Helper.CmToPx)
                Else
                    ScrollViewer.LineUp()
                End If
                e.Handled = True
            Case Key.PageUp
                If Keyboard.Modifiers = ModifierKeys.Control Then
                    ScrollViewer.PageLeft()
                Else
                    ScrollViewer.PageUp()
                End If
                e.Handled = True
            Case Key.Left
                If Keyboard.Modifiers = ModifierKeys.Control Then
                    ScrollViewer.ScrollToHorizontalOffset(ScrollViewer.HorizontalOffset - Helper.CmToPx)
                Else
                    ScrollViewer.LineLeft()
                End If
                e.Handled = True
            Case Key.Down
                If Keyboard.Modifiers = ModifierKeys.Control Then
                    ScrollViewer.ScrollToVerticalOffset(ScrollViewer.VerticalOffset + Helper.CmToPx)
                Else
                    ScrollViewer.LineDown()
                End If
                e.Handled = True
            Case Key.PageDown
                If Keyboard.Modifiers = ModifierKeys.Control Then
                    ScrollViewer.PageRight()
                Else
                    ScrollViewer.PageDown()
                End If
                e.Handled = True
            Case Key.Right
                If Keyboard.Modifiers = ModifierKeys.Control Then
                    ScrollViewer.ScrollToHorizontalOffset(ScrollViewer.HorizontalOffset + Helper.CmToPx)
                Else
                    ScrollViewer.LineRight()
                End If
                e.Handled = True
        End Select
    End Sub

    Dim SelStartPoint As Point

    Private Sub Designer_MouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs) Handles Me.MouseLeftButtonDown
        Me.CaptureMouse()
        SelStartPoint = e.GetPosition(Me.ConnectionCanvas)
        Canvas.SetLeft(SelectionBorder, SelStartPoint.X)
        Canvas.SetTop(SelectionBorder, SelStartPoint.Y)
        SelectionBorder.Width = 0
        SelectionBorder.Height = 0
        SelectionBorder.Visibility = Visibility.Visible
    End Sub

    Private Sub Designer_MouseMove(sender As Object, e As MouseEventArgs) Handles Me.MouseMove
        If e.LeftButton = MouseButtonState.Pressed Then
            Dim R As New Rect(SelStartPoint, e.GetPosition(Me.ConnectionCanvas))
            Canvas.SetLeft(SelectionBorder, R.Left)
            Canvas.SetTop(SelectionBorder, R.Top)
            SelectionBorder.Width = R.Width
            SelectionBorder.Height = R.Height
        End If

    End Sub

    Private Sub Designer_MouseLeftButtonUp(sender As Object, e As MouseButtonEventArgs) Handles Me.MouseLeftButtonUp
        If SelectionBorder.Visibility = Visibility.Visible Then
            SelectionBorder.Visibility = Visibility.Collapsed
            Dim R As New Rect(SelStartPoint, e.GetPosition(Me.ConnectionCanvas))
            If R.Width = 0 AndAlso R.Height = 0 Then Return

            If SelectedToolBoxItem Is Nothing Then
                VisualTreeHelper.HitTest(
                    DesignerCanvas, Nothing,
                    AddressOf GetDiagramsHitTestResult,
                    New GeometryHitTestParameters(New RectangleGeometry(R))
                )
            Else
                DarwDiagram(SelectedToolBoxItem, R.TopLeft, R.Width, R.Height)
            End If
        End If

        Me.ReleaseMouseCapture()
    End Sub

    Private Sub Designer_PreviewMouseWheel(sender As Object, e As MouseWheelEventArgs) Handles Me.PreviewMouseWheel
        If Keyboard.Modifiers = ModifierKeys.Control Then
            Dim Value As Integer = Me.Scale * 100 + (e.Delta / 24)
            If Not (Value < 25 OrElse Value > 500) Then Me.Scale = Value / 100
            e.Handled = True
        End If
    End Sub

    Private Function GetDiagramsHitTestResult(result As GeometryHitTestResult) As HitTestResultBehavior
        If Not result.IntersectionDetail = IntersectionDetail.Empty Then
            Dim ui = TryCast(result.VisualHit, UIElement)
            If ui IsNot Nothing AndAlso ui.IsVisible Then
                Dim Item = Helper.GetListBoxItem(ui)
                If Item IsNot Nothing Then
                    Item.IsSelected = True
                    Item.Focus()
                End If
            End If
        End If
        Return HitTestResultBehavior.Continue
    End Function

    Private Sub UndoStack_UndoRedoStateChanged(CanUndo As Boolean, CanRedo As Boolean) Handles UndoStack.UndoRedoStateChanged
        Me.HasChanges = CanUndo 'OrElse CanRedo
        Me.CanUndo = CanUndo
        Me.CanRedo = CanRedo
    End Sub

    Sub UpdateGrid()
        Helper.UpdateControl(Me.ScrollViewer)
        Dim VpWidth = Me.ScrollViewer.ViewportWidth
        Dim VpHeight = Me.ScrollViewer.ViewportHeight

        Dim W = Me.PageWidth * Me.Scale
        Dim H = Me.PageHeight * Me.Scale
        Dim x As Double = If(VpWidth > W, (VpWidth - W) / 2, 0)
        Dim y As Double = If(VpHeight > H, (VpHeight - H) / 2, 0)

        GridLinesBorder.Margin = New Thickness(x, y, 0, 0)
        GridLinesBorder.Width = Math.Min(VpWidth, W)
        GridLinesBorder.Height = Math.Min(VpHeight, H)

    End Sub

    Public Function GetControlName(index As Integer) As String
        If index = -1 Then Return Me.Name
        If index < Me.Items.Count Then Return GetControlName(Me.Items(index))
        Return ""
    End Function

    Public Function GetControlName(Optional control As UIElement = Nothing) As String
        If control Is Nothing Then control = Me.SelectedItem
        Dim controlName = Automation.AutomationProperties.GetName(control)
        If controlName = "" Then
            Dim fw = TryCast(control, FrameworkElement)
            If fw IsNot Nothing Then controlName = fw.Name
        End If
        Return controlName
    End Function

    Public Function GetControlText(Optional control As UIElement = Nothing) As String
        If control Is Nothing Then control = Me.SelectedItem

        Try
            Return CObj(control).Text
        Catch
            Try
                Dim contentControl = TryCast(control, ContentControl)
                If contentControl IsNot Nothing Then
                    Dim content = contentControl.Content
                    If content Is Nothing Then Return ""

                    If TypeOf content Is String Then
                        Return CStr(content)
                    Else
                        Return content.Text
                    End If
                End If
            Catch
            End Try
        End Try

        Dim txt = Automation.AutomationProperties.GetHelpText(control)
        Return If(txt, "")
    End Function

    Public Sub SetControlText(controlIndex As Integer, value As String)
        If controlIndex = -1 Then
            Automation.AutomationProperties.SetHelpText(Me, _Text)
            Dim OldState As New PropertyState(
                    Sub()
                        Me.SelectedIndex = -1
                        _Text = Automation.AutomationProperties.GetHelpText(Me)
                        RaiseEvent PageShown(UpdateControlNameAndText)
                    End Sub,
                    Me, Automation.AutomationProperties.HelpTextProperty)

            _Text = value
            Automation.AutomationProperties.SetHelpText(Me, value)
            UndoStack.ReportChanges(New UndoRedoUnit(OldState.SetNewValue()))

        Else
            SetControlText(Me.Items(controlIndex), value, True, True)
        End If
    End Sub

    Friend Sub SetControlText(control As UIElement, value As String, Optional trySetText As Boolean = True, Optional reportChanges As Boolean = False)
        If control Is Nothing Then control = Me.SelectedItem

        Dim txt = GetControlText(control)
        If txt = value Then Return

        If reportChanges Then
            Automation.AutomationProperties.SetHelpText(control, txt)
            Dim OldState As New PropertyState(
                    Sub()
                        Me.SelectedItem = control
                        SetControlText(control, Automation.AutomationProperties.GetHelpText(control))
                        RaiseEvent PageShown(UpdateControlNameAndText)
                    End Sub,
                    control, Automation.AutomationProperties.HelpTextProperty)

            Automation.AutomationProperties.SetHelpText(control, value)
            UndoStack.ReportChanges(New UndoRedoUnit(OldState.SetNewValue()))
        End If

        Dim contentControl = TryCast(control, ContentControl)
        If contentControl IsNot Nothing Then
            contentControl.Content = New TextBlock() With {
                .Text = value,
                .TextWrapping = TextWrapping.Wrap
            }

        ElseIf trySetText Then
            Try
                CObj(control).Text = value
            Catch
            End Try
        End If

        RaiseEvent PageShown(UpdateControlNameAndText) ' To Update text
    End Sub


    Public Function GetControlNameOrDefault(control As UIElement) As String
        Dim controlName = GetControlName(control)
        If controlName = "" Then
            Dim t = GetType(Control)
            controlName = GetDefaultControlName(t, t.Name)
            Automation.AutomationProperties.SetName(control, controlName)
        End If
        Return controlName
    End Function

    ' -------------------------------------------------------------------------------------------------
#Region "Dependancy Properties"

    Dim ExitChange As Boolean = False
    Friend MinZIndex As Integer

#Region "Left Attached Property"
    <TypeConverter(GetType(LengthConverter))>
    Public Shared Function GetLeft(element As DependencyObject) As Double
        If element Is Nothing Then
            Throw New ArgumentNullException("element")
        End If

        Return element.GetValue(LeftProperty)
    End Function

    <TypeConverter(GetType(LengthConverter))>
    Public Shared Sub SetLeft(element As DependencyObject, value As Double)
        If element Is Nothing Then Return

        element.SetValue(LeftProperty, value)
    End Sub

    Shared Sub LeftChanged(dpo As DependencyObject, e As DependencyPropertyChangedEventArgs)
        Dim d = Helper.GetDesigner(dpo)
        If d Is Nothing OrElse d.ExitChange Then Return

        Dim Item As ListBoxItem
        Dim Pnl As DiagramPanel
        Dim Diagram As UIElement
        Helper.GetTreeElemets(dpo, Item, Pnl, Diagram)
        d.ExitChange = True
        Designer.SetLeft(Item, e.NewValue)
        Canvas.SetLeft(Item, e.NewValue)
        Designer.SetLeft(Pnl, e.NewValue)
        Designer.SetLeft(Diagram, e.NewValue)
        d.ExitChange = False
    End Sub

    Public Shared ReadOnly LeftProperty As _
                           DependencyProperty = DependencyProperty.RegisterAttached("Left",
                           GetType(Double), GetType(Designer), New PropertyMetadata(AddressOf LeftChanged))
#End Region

#Region "Right Attached Property"
    <TypeConverter(GetType(LengthConverter))>
    Public Shared Function GetRight(element As DependencyObject) As Double
        If element Is Nothing Then
            Throw New ArgumentNullException("element")
        End If

        Return element.GetValue(RightProperty)
    End Function

    <TypeConverter(GetType(LengthConverter))>
    Public Shared Sub SetRight(element As DependencyObject, value As Double)
        If element Is Nothing Then
            Return
        End If

        element.SetValue(RightProperty, value)

    End Sub

    Shared Sub RightChanged(dpo As DependencyObject, e As DependencyPropertyChangedEventArgs)
        Dim d = Helper.GetDesigner(dpo)
        If d Is Nothing OrElse d.ExitChange Then Return

        Dim Item As ListBoxItem
        Dim Pnl As DiagramPanel
        Dim Diagram As UIElement

        Helper.GetTreeElemets(dpo, Item, Pnl, Diagram)

        d.ExitChange = True
        Designer.SetRight(Item, e.NewValue)
        Canvas.SetRight(Item, e.NewValue)
        Designer.SetRight(Pnl, e.NewValue)
        Designer.SetRight(Diagram, e.NewValue)
        d.ExitChange = False
    End Sub

    Public Shared ReadOnly RightProperty As _
                           DependencyProperty = DependencyProperty.RegisterAttached("Right",
                           GetType(Double), GetType(Designer), New PropertyMetadata(AddressOf RightChanged))
#End Region

#Region "Top Attached Property"
    <TypeConverter(GetType(LengthConverter))>
    Public Shared Function GetTop(element As DependencyObject) As Double
        If element Is Nothing Then
            Throw New ArgumentNullException("element")
        End If

        Return element.GetValue(TopProperty)
    End Function

    <TypeConverter(GetType(LengthConverter))>
    Public Shared Sub SetTop(element As DependencyObject, value As Double)
        If element Is Nothing Then
            Return
        End If

        element.SetValue(TopProperty, value)

    End Sub

    Shared Sub TopChanged(dpo As DependencyObject, e As DependencyPropertyChangedEventArgs)
        Dim d = Helper.GetDesigner(dpo)
        If d Is Nothing OrElse d.ExitChange Then Return

        Dim Item As ListBoxItem
        Dim Pnl As DiagramPanel
        Dim Diagram As UIElement

        Helper.GetTreeElemets(dpo, Item, Pnl, Diagram)

        d.ExitChange = True
        Designer.SetTop(Item, e.NewValue)
        Canvas.SetTop(Item, e.NewValue)
        Designer.SetTop(Pnl, e.NewValue)
        Designer.SetTop(Diagram, e.NewValue)
        d.ExitChange = False
    End Sub

    Public Shared ReadOnly TopProperty As _
                           DependencyProperty = DependencyProperty.RegisterAttached("Top",
                           GetType(Double), GetType(Designer), New PropertyMetadata(AddressOf TopChanged))
#End Region

#Region "Bottom Attached Property"
    <TypeConverter(GetType(LengthConverter))>
    Public Shared Function GetBottom(element As DependencyObject) As Double
        If element Is Nothing Then
            Throw New ArgumentNullException("element")
        End If

        Return element.GetValue(BottomProperty)
    End Function

    <TypeConverter(GetType(LengthConverter))>
    Public Shared Sub SetBottom(element As DependencyObject, value As Double)
        If element Is Nothing Then
            Return
        End If

        element.SetValue(BottomProperty, value)
    End Sub

    Shared Sub BottomChanged(dpo As DependencyObject, e As DependencyPropertyChangedEventArgs)
        Dim d = Helper.GetDesigner(dpo)
        If d Is Nothing OrElse d.ExitChange Then Return

        Dim Item As ListBoxItem
        Dim Pnl As DiagramPanel
        Dim Diagram As UIElement

        Helper.GetTreeElemets(dpo, Item, Pnl, Diagram)

        d.ExitChange = True
        Designer.SetBottom(Item, e.NewValue)
        Canvas.SetBottom(Item, e.NewValue)
        Designer.SetBottom(Pnl, e.NewValue)
        Designer.SetBottom(Diagram, e.NewValue)
        d.ExitChange = False
    End Sub

    Public Shared ReadOnly BottomProperty As _
                           DependencyProperty = DependencyProperty.RegisterAttached("Bottom",
                           GetType(Double), GetType(Designer), New PropertyMetadata(AddressOf BottomChanged))
#End Region

#Region "FrameWidth Attached Property"
    <TypeConverter(GetType(LengthConverter))>
    Public Shared Function GetFrameWidth(element As DependencyObject) As Double
        If element Is Nothing Then
            Throw New ArgumentNullException("element")
        End If

        Return element.GetValue(FrameWidthProperty)
    End Function

    <TypeConverter(GetType(LengthConverter))>
    Public Shared Sub SetFrameWidth(element As DependencyObject, value As Double)
        If element Is Nothing Then
            Return
        End If

        element.SetValue(FrameWidthProperty, value)
    End Sub

    Shared Sub FrameWidthChanged(Diagram As DependencyObject, e As DependencyPropertyChangedEventArgs)
        Dim Pnl = Helper.GetDiagramPanel(Diagram)
        If Pnl Is Nothing Then Return
        Pnl.Width = e.NewValue
    End Sub

    Public Shared ReadOnly FrameWidthProperty As _
                           DependencyProperty = DependencyProperty.RegisterAttached("FrameWidth",
                           GetType(Double), GetType(Designer), New PropertyMetadata(Double.NaN, AddressOf FrameWidthChanged))

#End Region

#Region "FrameHeight Attached Property"
    <TypeConverter(GetType(LengthConverter))>
    Public Shared Function GetFrameHeight(element As DependencyObject) As Double
        If element Is Nothing Then
            Throw New ArgumentNullException("element")
        End If

        Return element.GetValue(FrameHeightProperty)
    End Function

    <TypeConverter(GetType(LengthConverter))>
    Public Shared Sub SetFrameHeight(element As DependencyObject, value As Double)
        If element Is Nothing Then
            Return
        End If

        element.SetValue(FrameHeightProperty, value)
    End Sub

    Shared Sub FrameHeightChanged(Diagram As DependencyObject, e As DependencyPropertyChangedEventArgs)
        Dim Pnl = Helper.GetDiagramPanel(Diagram)
        If Pnl Is Nothing Then Return
        Pnl.Height = e.NewValue
    End Sub

    Public Shared ReadOnly FrameHeightProperty As _
                           DependencyProperty = DependencyProperty.RegisterAttached("FrameHeight",
                           GetType(Double), GetType(Designer), New PropertyMetadata(Double.NaN, AddressOf FrameHeightChanged))

#End Region

#Region "RotationAngle Attached Property"

    Public Shared Function GetRotationAngle(element As DependencyObject) As Double
        If element Is Nothing Then
            Throw New ArgumentNullException("element")
        End If

        Return element.GetValue(RotationAngleProperty)
    End Function

    Public Shared Sub SetRotationAngle(element As DependencyObject, value As Double)
        If element Is Nothing Then
            Return
        End If

        element.SetValue(RotationAngleProperty, value)
    End Sub

    Shared Sub RotationAngleChanged(Diagram As DependencyObject, e As DependencyPropertyChangedEventArgs)
        Dim Item = Helper.GetListBoxItem(Diagram)
        If Item Is Nothing Then Return
        Item.RenderTransform = New RotateTransform(e.NewValue)
    End Sub

    Public Shared ReadOnly RotationAngleProperty As _
                           DependencyProperty = DependencyProperty.RegisterAttached("RotationAngle",
                           GetType(Double), GetType(Designer), New PropertyMetadata(0.0, AddressOf RotationAngleChanged))

#End Region

#Region "ScrollViewer"

    Friend Property ScrollViewer As ScrollViewer
        Get
            Return GetValue(ScrollViewerProperty)
        End Get

        Set(value As ScrollViewer)
            SetValue(ScrollViewerProperty, value)
        End Set
    End Property

    Public Shared ReadOnly ScrollViewerProperty As DependencyProperty =
                           DependencyProperty.Register("ScrollViewer",
                           GetType(ScrollViewer), GetType(Designer),
                           New PropertyMetadata(Nothing))

#End Region

#Region "DesignerGrid"
    Friend Property DesignerGrid As Grid
        Get
            Return GetValue(DesignerGridProperty)
        End Get

        Private Set(value As Grid)
            SetValue(DesignerGridProperty, value)
        End Set
    End Property

    Public Shared ReadOnly DesignerGridProperty As DependencyProperty =
                           DependencyProperty.Register("DesignerGrid",
                           GetType(Grid), GetType(Designer),
                           New PropertyMetadata(Nothing))
#End Region

#Region "PageWidth"
    <TypeConverter(GetType(LengthConverter))>
    Public Property PageWidth As Double
        Get
            Return GetValue(PageWidthProperty)
        End Get

        Set(value As Double)
            SetValue(PageWidthProperty, value)
        End Set
    End Property

    Shared Sub PageWidthChanged(sender As System.Windows.DependencyObject, e As System.Windows.DependencyPropertyChangedEventArgs)
        Dim Designer As Designer = sender
        If Designer.DesignerCanvas Is Nothing Then Return
        Designer.DesignerCanvas.Width = e.NewValue
        Designer.ConnectionCanvas.Width = e.NewValue
        Designer.UpdateGrid()
    End Sub

    Public Shared ReadOnly PageWidthProperty As DependencyProperty =
                           DependencyProperty.Register("PageWidth",
                           GetType(Double), GetType(Designer),
                           New PropertyMetadata(21 * Helper.CmToPx, AddressOf PageWidthChanged))

#End Region

#Region "PageHeight"
    <TypeConverter(GetType(LengthConverter))>
    Public Property PageHeight As Double
        Get
            Return GetValue(PageHeightProperty)
        End Get

        Set(value As Double)
            SetValue(PageHeightProperty, value)
        End Set
    End Property

    Shared Sub PageHeightChanged(sender As System.Windows.DependencyObject, e As System.Windows.DependencyPropertyChangedEventArgs)
        Dim Designer As Designer = sender
        If Designer.DesignerCanvas Is Nothing Then Return
        Designer.DesignerCanvas.Height = e.NewValue
        Designer.ConnectionCanvas.Height = e.NewValue
        Designer.UpdateGrid()
    End Sub

    Public Shared ReadOnly PageHeightProperty As DependencyProperty =
                           DependencyProperty.Register("PageHeight",
                           GetType(Double), GetType(Designer),
                           New PropertyMetadata(29.7 * Helper.CmToPx, AddressOf PageHeightChanged))

#End Region

#Region "ShowGrid"


    Public Property ShowGrid As Boolean
        Get
            Return GetValue(ShowGridProperty)
        End Get

        Set(value As Boolean)
            SetValue(ShowGridProperty, value)
        End Set
    End Property

    Shared Sub ShowGridChanged(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
        Dim Designer As Designer = d
        If Designer.DesignerCanvas Is Nothing Then Return

        If e.NewValue Then
            Designer.GridLinesBorder.Visibility = Windows.Visibility.Visible
        Else
            Designer.GridLinesBorder.Visibility = Windows.Visibility.Collapsed
        End If
    End Sub

    Public Shared ReadOnly ShowGridProperty As DependencyProperty =
                           DependencyProperty.Register("ShowGrid",
                           GetType(Boolean), GetType(Designer),
                           New PropertyMetadata(False, AddressOf ShowGridChanged))


#End Region

#Region "Scale"
    Public Property Scale As Double
        Get
            Return GetValue(ScaleProperty)
        End Get

        Set(value As Double)
            SetValue(ScaleProperty, value)
        End Set
    End Property

    Shared Sub ScaleChanged(Designer As Designer, e As DependencyPropertyChangedEventArgs)
        Dim Sc As New ScaleTransform(e.NewValue, e.NewValue)
        Designer.DesignerCanvas.LayoutTransform = Sc
        Designer.ConnectionCanvas.LayoutTransform = Sc
        Designer.GridBrush.Transform = Sc
        Designer.TbLeftLocation.LayoutTransform = Sc
        Designer.TbTopLocation.LayoutTransform = Sc

        If Designer.ZoomBox IsNot Nothing Then Designer.ZoomBox.zoomSlider.Value = e.NewValue * 100

        Designer.UpdateGrid()
    End Sub

    Shared Function VlaidateScale(Value As Double) As Boolean
        Return (Value > 0 AndAlso Value <= 5)
    End Function

    Public Shared ReadOnly ScaleProperty As DependencyProperty =
                           DependencyProperty.Register("Scale",
                           GetType(Double), GetType(Designer),
                           New PropertyMetadata(1.0, AddressOf ScaleChanged), AddressOf VlaidateScale)

#End Region



#Region "DiagramTextFontProps"
    Public Shared Function GetDiagramTextFontProps(element As DependencyObject) As PropertyDictionary
        If element Is Nothing Then
            Throw New ArgumentNullException("element")
        End If

        Return element.GetValue(DiagramTextFontPropsProperty)
    End Function

    Public Shared Sub SetDiagramTextFontProps(element As DependencyObject, value As PropertyDictionary)
        If element Is Nothing Then
            Throw New ArgumentNullException("element")
        End If

        element.SetValue(DiagramTextFontPropsProperty, value)
    End Sub

    Public Shared ReadOnly DiagramTextFontPropsProperty As _
                           DependencyProperty = DependencyProperty.RegisterAttached("DiagramTextFontProps",
                           GetType(PropertyDictionary), GetType(Designer),
                           New PropertyMetadata(AddressOf DiagramTextFontPropsChanged))

    Shared Sub DiagramTextFontPropsChanged(Diagram As DependencyObject, e As DependencyPropertyChangedEventArgs)
        Dim FontProps As PropertyDictionary = e.NewValue
        If FontProps Is Nothing Then Return

        FontProps.SetDependencyObject(Diagram)
        FontProps.SetValuesToObj()
    End Sub

#End Region


#Region "HasChanges"

    Public Property HasChanges As Boolean
        Get
            Return GetValue(HasChangesProperty)
        End Get

        Set(value As Boolean)
            SetValue(HasChangesProperty, value)
        End Set
    End Property

    Public Shared ReadOnly HasChangesProperty As DependencyProperty =
                           DependencyProperty.Register("HasChanges",
                           GetType(Boolean), GetType(Designer),
                           New PropertyMetadata(False))

#End Region

#Region "CanUndo"

    Public Property CanUndo As Boolean
        Get
            Return GetValue(CanUndoProperty)
        End Get

        Set(value As Boolean)
            SetValue(CanUndoProperty, value)
        End Set
    End Property

    Public Shared ReadOnly CanUndoProperty As DependencyProperty =
                           DependencyProperty.Register("CanUndo",
                           GetType(Boolean), GetType(Designer),
                           New PropertyMetadata(False))

#End Region

#Region "CanRedo"

    Public Property CanRedo As Boolean
        Get
            Return GetValue(CanRedoProperty)
        End Get

        Set(value As Boolean)
            SetValue(CanRedoProperty, value)
        End Set
    End Property

    Public Property Text As String

    Private ReadOnly Property IsDirty As Boolean
        Get
            If IO.Path.GetFileNameWithoutExtension(_xamlFile).ToLower() = "global" Then Return False
            Return Me.HasChanges OrElse (
                _xamlFile = "" AndAlso _codeFile <> "" AndAlso
                Not _codeFile.ToLower().StartsWith(TempProjectPath)
            )
        End Get
    End Property

    Friend Shared ReadOnly Property PageCount As Integer
        Get
            Return FormKeys.Count
        End Get
    End Property

    Public Shared ReadOnly CanRedoProperty As DependencyProperty =
                           DependencyProperty.Register("CanRedo",
                           GetType(Boolean), GetType(Designer),
                           New PropertyMetadata(False))

#End Region

#Region "GroupID"

    Public Shared Function GetGroupID(element As DependencyObject) As Long
        If element Is Nothing Then
            Throw New ArgumentNullException("element")
        End If

        Return element.GetValue(GroupIDProperty)
    End Function

    Public Shared Sub SetGroupID(element As DependencyObject, value As Long)
        If element Is Nothing Then
            Throw New ArgumentNullException("element")
        End If

        element.SetValue(GroupIDProperty, value)

    End Sub

    Public Shared ReadOnly GroupIDProperty As _
                           DependencyProperty = DependencyProperty.RegisterAttached("GroupID",
                           GetType(Long), GetType(Designer),
                           New PropertyMetadata(AddressOf GroupIDChanged))

    Public Const UpdateFormName As Integer = -2
    Public Const UpdateControlNameAndText As Integer = -3
    Public Const GlobalFileIndex As Integer = -4

    Shared Sub GroupIDChanged(Diagram As DependencyObject, e As DependencyPropertyChangedEventArgs)
        Dim Pnl = Helper.GetDiagramPanel(Diagram)
        If Pnl Is Nothing Then Return
        DiagramGroup.Add(Pnl, CLng(e.NewValue))
    End Sub

#End Region


#End Region




End Class
