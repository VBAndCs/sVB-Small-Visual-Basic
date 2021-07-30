Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.Windows.Controls.Primitives
Imports System.Windows.Markup
Imports System.Xml

Public Class Designer
    Inherits ListBox

    Friend GridLinesBorder As Border
    Friend ConnectionCanvas As Canvas
    Friend DesignerCanvas As Canvas
    Friend TbTopLocation As TextBlock
    Friend TbLeftLocation As TextBlock
    Friend GridBrush As DrawingBrush
    Friend Editor As TextBox
    Friend ZoomBox As ZoomBox
    Friend MaxZIndex As Integer = 0
    Friend Shared Editing As Boolean = False
    Friend SelectionBorder As Border
    Friend WithEvents UndoStack As New UndoRedoStack(Of UndoRedoUnit)(1000)
    Dim DeleteUndoUnit As UndoRedoUnit
    Dim ConnectionsOldState As ConnectionState
    Friend SelectedBounds As Rect
    Friend GridPen As Pen

    Dim _fileName As String = ""
    Public Property FileName As String
        Get
            Return _fileName
        End Get

        Set(value As String)
            _fileName = If(value = "", "", IO.Path.GetFullPath(value))
            If _fileName <> "" AndAlso _CodeFilePath = "" Then
                _CodeFilePath = _fileName.Substring(0, _fileName.Length - 5) & ".sb"
            End If
        End Set
    End Property

    Public Sub New()
        Dim resourceLocater As Uri = New Uri("/DiagramHelper;component/Resources/designerdecorator.xaml", System.UriKind.Relative)
        Dim ResDec As ResourceDictionary = Application.LoadComponent(resourceLocater)
        Me.Resources.MergedDictionaries.Add(ResDec)
        Me.Style = FindResource("CanvasListBoxStyle")
        Me.ItemContainerStyle = FindResource("listBoxItemStyle")
        Me.SelectionMode = Controls.SelectionMode.Multiple
        Me.AllowDrop = True
    End Sub

    Private Shared NewPageOpened As Boolean

    Private Sub Designer_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        Dim Brdr As Border = VisualTreeHelper.GetChild(Me, 0)
        Dim G = CType(Brdr.Child, Grid)
        Me.ScrollViewer = G.Children(0)
        GridLinesBorder = G.Children(1)
        G = ScrollViewer.Content
        Me.DesignerCanvas = VisualTreeHelper.GetChild(G.Children(0), 0)
        DesignerGrid = G
        Me.ConnectionCanvas = G.Children(1)
        Me.DesignerCanvas.Width = Me.PageWidth
        Me.DesignerCanvas.Height = Me.PageHeight
        Me.ConnectionCanvas.Width = Me.PageWidth
        Me.ConnectionCanvas.Height = Me.PageHeight
        Me.DesignerCanvas.Background = SystemColors.ControlBrush
        Me.SelectionBorder = ConnectionCanvas.Children(0)

        ShowGridChanged(Me, New DependencyPropertyChangedEventArgs(ShowGridProperty, False, Me.ShowGrid))
        GridBrush = Me.FindResource("GridBrush")
        GridPen = FindResource("GridLinesPen")

        AddHandler Me.ScrollViewer.ScrollChanged, AddressOf ScrollViewer_ScrollChanged

        TbTopLocation = Me.Template.FindName("PART_TopLocation", Me)
        TbLeftLocation = Me.Template.FindName("PART_LeftLocation", Me)
        Editor = Me.Template.FindName("PART_Editor", Me)

        If Not NewPageOpened Then
            CurrentPage = Me
            Dispatcher.Invoke(Sub() OpenNewPage(False))
            NewPageOpened = True
        End If
        Me.Focus()

    End Sub

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
        If Me.GetControlName(control).ToLower() = newName Then
            Return True
        End If

        If IsNumeric(newName(0)) Then
            MessageBox.Show("Name can't start with a number.", "Invalid Name", MessageBoxButton.OK, MessageBoxImage.Error)
            Return False
        End If

        For Each cnt In Me.Items
            If Me.GetControlName(cnt).ToLower() = newName Then
                MessageBox.Show($"There is another control named '{newName}'!", "Invalid Name", MessageBoxButton.OK, MessageBoxImage.Error)
                Return False
            End If
        Next

        Try
            Dim OldState As New PropertyState(
                Sub()
                    Me.SelectedItem = control
                    SetControlName(control, Automation.AutomationProperties.GetName(control))
                    RaiseEvent PageShown(-3)
                End Sub,
                control,
                Automation.AutomationProperties.NameProperty)

            Dim fw = TryCast(control, FrameworkElement)
            If fw IsNot Nothing Then fw.Name = name
            Automation.AutomationProperties.SetName(control, name)
            Me.UndoStack.ReportChanges(New UndoRedoUnit(OldState.SetNewValue))
            RaiseEvent PageShown(-3)
            Return True
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Invalid Name", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try

        Return False
    End Function

    Public Function ChangeFormName(name As String) As Boolean
        If name = "" Then Return False

        Dim newName = name.ToLower()
        If Me.Name.ToLower() = newName Then Return True

        If IsNumeric(newName(0)) Then
            MessageBox.Show("Name can't start with a number.", "Invalid Name", MessageBoxButton.OK, MessageBoxImage.Error)
            Return False
        End If

        For Each page In Pages.Values
            If page.Name.ToLower() = newName Then
                MessageBox.Show($"There is another form named '{newName}'!", "Invalid Name", MessageBoxButton.OK, MessageBoxImage.Error)
                Return False
            End If
        Next

        Try
            Dim OldState As New PropertyState(AddressOf UpdateFormInfo, Me, Designer.NameProperty)
            Me.Name = name
            Me.UndoStack.ReportChanges(New UndoRedoUnit(OldState.SetNewValue))
            UpdateFormInfo()
            RaiseEvent PageShown(-2)
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

#Region "Connections"

    Friend Connections As New Dictionary(Of UIElement, List(Of Connection))
    Friend ConnectionMode As Boolean
    Friend ConnectionSourceDiagram As FrameworkElement
    Friend ConnectionTargetDiagram As FrameworkElement
    Friend SourceConnector As ConnectorThumb
    Friend TargetConnector As ConnectorThumb

    Public Function GetConnectionsFromSource(Diagram As UIElement) As List(Of Connection)
        If Connections.ContainsKey(Diagram) Then Return Connections(Diagram)
        Dim Lst As New List(Of Connection)
        Connections.Add(Diagram, Lst)
        Return Lst
    End Function

    Public Function GetConnectionsToTarget(Diagram As UIElement) As List(Of Connection)
        Dim Conns = From Lst In Connections.Values
                    From C In Lst
                    Where C.TargetDiagram Is Diagram
                    Select C

        Return Conns.ToList
    End Function

    Public Function GetConnection(Diagram1 As UIElement, Diagram2 As UIElement) As Connection
        If Connections.ContainsKey(Diagram1) Then
            Dim Conns = From C In Connections(Diagram1) Where C.TargetDiagram Is Diagram2
            If Conns.Count > 0 Then Return Conns.First
        End If

        If Connections.ContainsKey(Diagram2) Then
            Dim Conns = From C In Connections(Diagram2) Where C.TargetDiagram Is Diagram1
            If Conns.Count > 0 Then Return Conns.First
        End If

        Return Nothing
    End Function

    Public Sub RemoveConnections(Diagram As UIElement)
        If Connections.ContainsKey(Diagram) Then
            Dim OutConns = Connections(Diagram)
            For i = OutConns.Count - 1 To 0 Step -1
                If ConnectionsOldState Is Nothing Then ConnectionsOldState = New ConnectionState(ConnectionAction.Create)
                ConnectionsOldState.Add(OutConns(i))
                OutConns(i).Delete()
            Next
        End If

        Dim Conns = From Lst In Connections.Values
                    From C In Lst
                    Where C.TargetDiagram Is Diagram
                    Select C

        For i = Conns.Count - 1 To 0 Step -1
            ConnectionsOldState.Add(Conns(i))
            Conns(i).Delete()
        Next
    End Sub

    Public Sub RemoveSelectedConnections()
        For Each ConList In Me.Connections.Values
            For i = ConList.Count - 1 To 0 Step -1
                Dim Con = ConList(i)
                If Con.IsSelected Then
                    ConnectionsOldState.Add(Con)
                    Con.Delete()
                End If
            Next
        Next
    End Sub

    Public Sub ClearConnections()
        For Each ConList In Connections.Values
            For i = ConList.Count - 1 To 0 Step -1
                ConList(i).Delete()
            Next
        Next
    End Sub

#End Region

#Region "Pages"
    Public Shared Pages As New Dictionary(Of String, Designer)
    Public Shared FormNames As New ObservableCollection(Of String)
    Public Shared FormKeys As New List(Of String)
    Public Shared CurrentPage As Designer
    Public Shared Property PagesGrid As Grid
    Public Property CodeFilePath As String = ""

    Private Shared TempKeyNum As Integer = 0
    Public PageKey As String

    Public Shared Event PageShown(index As Integer)

    Public ReadOnly Property IsNew As Boolean
        Get
            Return Not HasChanges AndAlso _fileName = ""
        End Get
    End Property


    Public Shared Function GetTempFormName() As String
        TempKeyNum += 1
        Return "Form" & TempKeyNum
    End Function

    Private Shared Function GetTempKey(Optional pagePath As String = "") As String
        If pagePath <> "" Then
            For Each p In Pages
                If p.Value._fileName = pagePath Then Return p.Key
            Next
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
        Dim displayName = " ● " & CurrentPage.Name & If(CurrentPage._fileName = "", " *", ".xaml")
        Dim i = FormKeys.IndexOf(CurrentPage.PageKey)
        If i = -1 Then
            FormNames.Add(displayName)
            FormKeys.Add(CurrentPage.PageKey)
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
        BringToFront()
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

    Public Shared Function ClosePage(Optional openNewPageIfLast As Boolean = True) As Boolean
        If CurrentPage.IsNew AndAlso Pages.Count = 1 Then Return True

        If CurrentPage.HasChanges OrElse (CurrentPage._fileName = "" AndAlso CurrentPage.CodeFilePath <> "") Then
            If Not CurrentPage.AskToSave() Then Return False
        End If

        Pages.Remove(CurrentPage.PageKey)
        PagesGrid.Children.Remove(CurrentPage)

        Dim i = FormKeys.IndexOf(CurrentPage.PageKey)
        If i <> -1 Then
            FormKeys.RemoveAt(i)
            FormNames.RemoveAt(i)
        End If

        Dim nextKey As String
        If Pages.Count > 0 Then
            Dim n = Pages.Count - 1
            If i > n Then
                nextKey = FormKeys(n)
            Else
                nextKey = FormKeys(i)
            End If

            SwitchTo(nextKey, False)

        ElseIf openNewPageIfLast Then
            OpenNewPage(False)
        End If

        Return True
    End Function

    Public Shared Function SwitchTo(key As String, Optional UpdateCurrentPage As Boolean = True) As String
        If key = "" Then Return OpenNewPage()

        If CurrentPage IsNot Nothing AndAlso UpdateCurrentPage Then UpdatePageInfo()

        If Pages.ContainsKey(key) Then
            CurrentPage = Pages(key)
        Else
            Dim xamlPath = key.ToLower()
            Dim codePath = xamlPath.Substring(0, xamlPath.Length - 5) & ".sb"
            For Each item In Pages
                Dim page = item.Value
                If page._fileName.ToLower() = xamlPath OrElse
                         page._CodeFilePath.ToLower() = codePath Then

                    ' File is already opened. Stwich to it.
                    Return SwitchTo(item.Key, False)
                End If
            Next

            ' File is not already opened. Open it.
            Open(key)
        End If

        BringToFront()
        RaiseEvent PageShown(FormKeys.IndexOf(CurrentPage.PageKey))
        Return CurrentPage.PageKey
    End Function

    Private Shared Sub BringToFront()
        Dim z = 0
        For Each item In PagesGrid.Children
            z = Math.Max(z, Grid.GetZIndex(item))
        Next

        Grid.SetZIndex(CurrentPage, z + 1)
    End Sub

    Private Function PageToXaml() As String
        For Each Diagram In Me.Items
            Commands.UpdateFontProperties(Diagram)
        Next

        Dim canvas As New Canvas
        canvas.Name = If(Me.Name = "", Me.PageKey.Replace("KEY", "Form"), Me.Name)
        canvas.Width = Me.PageWidth
        canvas.Height = Me.PageHeight
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
            diagram2.ClearValue(DiagramTextBlockPropertyKey)

            ' GroupID, DiagramText will cause isuues in Small Basic, because they need a reference to DiagramHelper
            ' Use any wpf built-in properety to hold theor values
            Dim txt = GetDiagramText(diagram2)
            If txt <> "" Then
                Automation.AutomationProperties.SetHelpText(diagram2, txt)
                diagram2.ClearValue(DiagramTextProperty)
            Else
                diagram2.ClearValue(Automation.AutomationProperties.HelpTextProperty)
            End If

            Dim gID = GetGroupID(diagram2)
            If gID <> "" Then
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



    Private Sub XamlToPage(xaml As String)
        Dim canvas As Canvas = XamlReader.Load(XmlReader.Create(New IO.StringReader(xaml)))
        Me.Name = canvas.Name
        _Text = Automation.AutomationProperties.GetHelpText(canvas)

        Me.Visibility = Visibility.Hidden

        If Not Double.IsNaN(canvas.Width) Then Me.PageWidth = canvas.Width

        If Not Double.IsNaN(canvas.Height) Then Me.PageHeight = canvas.Height

        For Each child In canvas.Children
            Dim Diagram = TryCast(Helper.Clone(child), FrameworkElement)
            If Diagram Is Nothing Then Continue For

            Me.Items.Add(Diagram)

            ' Restore the GroupID and DiagramText properties
            Dim txt = Automation.AutomationProperties.GetHelpText(Diagram)
            If txt = "" Then
                Diagram.ClearValue(DiagramTextProperty)
            Else
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
            Diagram.ClearValue(FrameworkElement.WidthProperty)

            Designer.SetFrameHeight(Diagram, Diagram.Height)
            Diagram.ClearValue(FrameworkElement.HeightProperty)

            Designer.SetLeft(Diagram, Canvas.GetLeft(Diagram))
            Designer.SetTop(Diagram, Canvas.GetTop(Diagram))

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
        Dim TbItem As ToolBoxItem = e.Data.GetData(GetType(ToolBoxItem))
        If TbItem IsNot Nothing Then
            Dim newItem = TbItem.Child
            Dim diagram As UIElement
            Dim defaultName = ""
            Dim controlType As Type
            Dim typeName As String
            Dim sbControl = TryCast(newItem, SBControl)

            If sbControl Is Nothing Then
                diagram = Helper.Clone(newItem)
                controlType = GetType(Label)
                typeName = TbItem.Name
            Else
                diagram = Helper.Clone(sbControl.Control)
                controlType = diagram.GetType()
                typeName = controlType.Name
            End If

            defaultName = GetDefaultControlName(controlType, typeName)

            If defaultName <> "" Then
                Automation.AutomationProperties.SetName(diagram, defaultName)
                SetControlText(diagram, defaultName, False)
            End If

            diagram.ClearValue(ToolTipProperty)

            Dim Pos = e.GetPosition(Me.DesignerCanvas)
            Dim OldState = New CollectionState(AddressOf AfterRestoreAction, Me.Items, diagram)
            Me.Items.Add(diagram)
            UndoStack.ReportChanges(New UndoRedoUnit(OldState.SetNewValue))

            Designer.SetLeft(diagram, Pos.X)
            Designer.SetTop(diagram, Pos.Y)
            Helper.UpdateControl(Me)
            Dim Item = Helper.GetListBoxItem(diagram)
            Me.SelectedIndex = -1
            Connection.DeselectAll(Me)
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
        RemoveConnections(Diagram)
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
        ConnectionsOldState = New ConnectionState(ConnectionAction.Create)
        Me.RemoveSelectedConnections()
        Me.RemoveSelectedDiagrams()

        If ConnectionsOldState.Count > 0 Then DeleteUndoUnit.Add(ConnectionsOldState)
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
        ConnectionsOldState = New ConnectionState(ConnectionAction.Create)
        Me.Copy()
        Me.RemoveSelectedDiagrams()

        If ConnectionsOldState.Count > 0 Then DeleteUndoUnit.Add(ConnectionsOldState)
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
        If _fileName = "" Then
            Return SaveAs()
        ElseIf Me.HasChanges Then
            Return SavePage("")
        Else
            Return True
        End If
    End Function

    Public Function SaveAs() As Boolean
        ' Configure open file dialog box 
        Dim dlg As New Microsoft.Win32.SaveFileDialog()
        dlg.DefaultExt = ".xaml" ' Default file extension
        dlg.Filter = "Diagram Pages|*.xaml"
        dlg.Title = "Save Diagram Design Page"

        Dim saveName As String
        If _fileName = "" Then
            saveName = Me.Name
        Else
            dlg.InitialDirectory = _fileName
            saveName = IO.Path.GetFileNameWithoutExtension(_fileName)
        End If
        dlg.FileName = saveName

        Dim result? As Boolean = dlg.ShowDialog()

        If result = True Then
            ' Don't use _fileName here, to update Tempkey!
            Me.FileName = dlg.FileName
            If Not SavePage("") Then Return False
            Return True
        End If
        Return False
    End Function

    Public Delegate Function SavePageDelegate(tmpPath As String) As Boolean
    Public SavePage As SavePageDelegate = AddressOf DoSave

    Public Function DoSave(Optional tmpPath As String = Nothing) As Boolean
        Try
            Dim xmal = PageToXaml()
            Dim saveTo = If(tmpPath = "", _fileName, tmpPath)
            IO.File.WriteAllText(saveTo, xmal, System.Text.Encoding.UTF8)
            _CodeFilePath = saveTo.Substring(0, saveTo.Length - 5) & ".sb"
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
        ImgSaver.Save(Me.DesignerGrid, _fileName)
        Me.Scale = Sc
    End Sub

    Public Sub Open()
        ' Configure open file dialog box 
        Dim dlg As New Microsoft.Win32.OpenFileDialog()
        dlg.DefaultExt = ".xaml" ' Default file extension
        dlg.Filter = "Diagram Pages|*.xaml"
        dlg.Title = "Open Diagram Design Page"
        If dlg.ShowDialog() = True Then
            SwitchTo(dlg.FileName)
            Me.HasChanges = False
        End If
    End Sub


    Public Shared Sub Open(fileName As String)
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
            CurrentPage._fileName = IO.Path.GetFullPath(fileName)
            CurrentPage.PageKey = GetTempKey(CurrentPage._fileName)
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
        If dialog.ShowDialog() = True Then dialog.PrintVisual(Me.DesignerGrid, _fileName)
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
        Connection.SelectAll(Me)
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
        If Not Editing AndAlso e.Source Is Me AndAlso Keyboard.Modifiers <> ModifierKeys.Control Then
            Me.SelectedIndex = -1
            Connection.DeselectAll(Me)
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
                                            Helper.CmToPx, Helper.CmToPx)

    End Sub

    Private Sub Designer_PreviewKeyDown(sender As Object, e As KeyEventArgs) Handles Me.PreviewKeyDown
        If e.OriginalSource Is Editor Then Return

        If e.Key = Input.Key.Escape Then
            If ConnectionMode Then
                SourceConnector.CancelDrawConnection()
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
                    Me.ClosePage()
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
        If Not ConnectionMode Then
            Me.CaptureMouse()
            SelStartPoint = e.GetPosition(Me.ConnectionCanvas)
            Canvas.SetLeft(SelectionBorder, SelStartPoint.X)
            Canvas.SetTop(SelectionBorder, SelStartPoint.Y)
            SelectionBorder.Width = 0
            SelectionBorder.Height = 0
            SelectionBorder.Visibility = Windows.Visibility.Visible
        End If
    End Sub

    Private Sub Designer_MouseMove(sender As Object, e As MouseEventArgs) Handles Me.MouseMove
        If ConnectionMode Then
            If SourceConnector IsNot Nothing Then SourceConnector.UpdateConnection()
        ElseIf e.LeftButton = MouseButtonState.Pressed Then
            Dim R As New Rect(SelStartPoint, e.GetPosition(Me.ConnectionCanvas))
            Canvas.SetLeft(SelectionBorder, R.Left)
            Canvas.SetTop(SelectionBorder, R.Top)
            SelectionBorder.Width = R.Width
            SelectionBorder.Height = R.Height
        End If

    End Sub

    Private Sub Designer_MouseLeftButtonUp(sender As Object, e As MouseButtonEventArgs) Handles Me.MouseLeftButtonUp
        If SelectionBorder.Visibility = Windows.Visibility.Visible Then
            SelectionBorder.Visibility = Windows.Visibility.Collapsed
            Dim R As New Rect(SelStartPoint, e.GetPosition(Me.ConnectionCanvas))
            If R.Width = 0 AndAlso R.Height = 0 Then Return

            VisualTreeHelper.HitTest(
                    DesignerCanvas, Nothing,
                    AddressOf GetDiagramsHitTestResult,
                    New GeometryHitTestParameters(New RectangleGeometry(R))
            )

            Connection.SelectIntersection(Me, R)
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
                Dim x = CObj(control).Content
                If x IsNot Nothing AndAlso TypeOf x Is String Then
                    Return CStr(CObj(control).Content)
                End If
            Catch
            End Try
        End Try

        Dim txt = Automation.AutomationProperties.GetHelpText(control)
        If txt = "" Then txt = GetDiagramText(control)
        Return If(txt, "")
    End Function

    Public Sub SetControlText(controlIndex As Integer, value As String)
        If controlIndex = -1 Then
            Automation.AutomationProperties.SetHelpText(Me, _Text)
            Dim OldState As New PropertyState(
                    Sub()
                        Me.SelectedIndex = -1
                        _Text = Automation.AutomationProperties.GetHelpText(Me)
                        RaiseEvent PageShown(-3)
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
                        RaiseEvent PageShown(-3)
                    End Sub,
                    control, Automation.AutomationProperties.HelpTextProperty)

            Automation.AutomationProperties.SetHelpText(control, value)
            UndoStack.ReportChanges(New UndoRedoUnit(OldState.SetNewValue()))
        End If

        If trySetText Then
            Try
                CObj(control).Text = value
                control.ClearValue(DiagramTextProperty)
                RaiseEvent PageShown(-3) ' To Update text
                Return
            Catch
            End Try
        End If

        Try
            Dim x = CObj(control).Content
            If x Is Nothing OrElse TypeOf x Is String Then
                CObj(control).Content = value
            End If
        Catch
            If trySetText Then SetDiagramText(control, value)
        End Try
        RaiseEvent PageShown(-3) ' To Update text
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

#Region "Left Attached Property"
    <TypeConverter(GetType(LengthConverter))>
    Public Shared Function GetLeft(ByVal element As DependencyObject) As Double
        If element Is Nothing Then
            Throw New ArgumentNullException("element")
        End If

        Return element.GetValue(LeftProperty)
    End Function

    <TypeConverter(GetType(LengthConverter))>
    Public Shared Sub SetLeft(ByVal element As DependencyObject, ByVal value As Double)
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
    Public Shared Function GetRight(ByVal element As DependencyObject) As Double
        If element Is Nothing Then
            Throw New ArgumentNullException("element")
        End If

        Return element.GetValue(RightProperty)
    End Function

    <TypeConverter(GetType(LengthConverter))>
    Public Shared Sub SetRight(ByVal element As DependencyObject, ByVal value As Double)
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
    Public Shared Function GetTop(ByVal element As DependencyObject) As Double
        If element Is Nothing Then
            Throw New ArgumentNullException("element")
        End If

        Return element.GetValue(TopProperty)
    End Function

    <TypeConverter(GetType(LengthConverter))>
    Public Shared Sub SetTop(ByVal element As DependencyObject, ByVal value As Double)
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
    Public Shared Function GetBottom(ByVal element As DependencyObject) As Double
        If element Is Nothing Then
            Throw New ArgumentNullException("element")
        End If

        Return element.GetValue(BottomProperty)
    End Function

    <TypeConverter(GetType(LengthConverter))>
    Public Shared Sub SetBottom(ByVal element As DependencyObject, ByVal value As Double)
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
    Public Shared Function GetFrameWidth(ByVal element As DependencyObject) As Double
        If element Is Nothing Then
            Throw New ArgumentNullException("element")
        End If

        Return element.GetValue(FrameWidthProperty)
    End Function

    <TypeConverter(GetType(LengthConverter))>
    Public Shared Sub SetFrameWidth(ByVal element As DependencyObject, ByVal value As Double)
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
    Public Shared Function GetFrameHeight(ByVal element As DependencyObject) As Double
        If element Is Nothing Then
            Throw New ArgumentNullException("element")
        End If

        Return element.GetValue(FrameHeightProperty)
    End Function

    <TypeConverter(GetType(LengthConverter))>
    Public Shared Sub SetFrameHeight(ByVal element As DependencyObject, ByVal value As Double)
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

    Public Shared Function GetRotationAngle(ByVal element As DependencyObject) As Double
        If element Is Nothing Then
            Throw New ArgumentNullException("element")
        End If

        Return element.GetValue(RotationAngleProperty)
    End Function

    Public Shared Sub SetRotationAngle(ByVal element As DependencyObject, ByVal value As Double)
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

        Set(ByVal value As ScrollViewer)
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

        Private Set(ByVal value As Grid)
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

        Set(ByVal value As Double)
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

        Set(ByVal value As Double)
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

        Set(ByVal value As Boolean)
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

        Set(ByVal value As Double)
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

#Region "DiagramTextBlock"

    Public Shared Function GetDiagramTextBlock(ByVal element As DependencyObject) As TextBlock
        If element Is Nothing Then
            Throw New ArgumentNullException("element")
        End If

        Return element.GetValue(DiagramTextBlockProperty)
    End Function

    Friend Shared Sub SetDiagramTextBlock(ByVal element As DependencyObject, value As TextBlock)
        If element Is Nothing Then
            Throw New ArgumentNullException("element")
        End If

        element.SetValue(DiagramTextBlockPropertyKey, value)
    End Sub

    Private Shared ReadOnly DiagramTextBlockPropertyKey As DependencyPropertyKey =
                        DependencyProperty.RegisterAttachedReadOnly("DiagramTextBlock",
                        GetType(TextBlock), GetType(Designer), New PropertyMetadata())

    Public Shared ReadOnly DiagramTextBlockProperty = DiagramTextBlockPropertyKey.DependencyProperty


#End Region

#Region "DiagramText"
    Public Shared Function GetDiagramText(ByVal element As DependencyObject) As String
        If element Is Nothing Then
            Throw New ArgumentNullException("element")
        End If

        Return element.GetValue(DiagramTextProperty)
    End Function

    Public Shared Sub SetDiagramText(ByVal element As DependencyObject, ByVal value As String)
        If element Is Nothing Then
            Throw New ArgumentNullException("element")
        End If

        element.SetValue(DiagramTextProperty, value)
    End Sub

    Public Shared ReadOnly DiagramTextProperty As _
                           DependencyProperty = DependencyProperty.RegisterAttached("DiagramText",
                           GetType(String), GetType(Designer),
                           New PropertyMetadata(AddressOf DiagramTextChanged))

    Shared Sub DiagramTextChanged(Diagram As DependencyObject, e As DependencyPropertyChangedEventArgs)
        Dim Pnl = Helper.GetDiagramPanel(Diagram)
        If Pnl Is Nothing Then Return
        Pnl.DiagramTextBlock.Text = e.NewValue
        If Pnl.DrawOutlineMenuItem.IsChecked Then Pnl.DiagramObj.OutlineText()
    End Sub

#End Region

#Region "DiagramTextFontProps"
    Public Shared Function GetDiagramTextFontProps(ByVal element As DependencyObject) As PropertyDictionary
        If element Is Nothing Then
            Throw New ArgumentNullException("element")
        End If

        Return element.GetValue(DiagramTextFontPropsProperty)
    End Function

    Public Shared Sub SetDiagramTextFontProps(ByVal element As DependencyObject, ByVal value As PropertyDictionary)
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
        Dim Pnl = Helper.GetDiagramPanel(Diagram)
        If Pnl Is Nothing Then Return

        Dim FontProps As PropertyDictionary = e.NewValue
        If FontProps Is Nothing Then Return

        FontProps.SetDependencyObject(Pnl.DiagramTextBlock)
        FontProps.SetValuesToObj()

        If Pnl.DrawOutlineMenuItem.IsChecked Then Pnl.DiagramObj.OutlineText()
    End Sub

#End Region

#Region "DiagramTextBackground"

    Public Shared Function GetDiagramTextBackground(ByVal element As DependencyObject) As Brush
        If element Is Nothing Then
            Throw New ArgumentNullException("element")
        End If

        Return element.GetValue(DiagramTextBackgroundProperty)
    End Function

    Public Shared Sub SetDiagramTextBackground(ByVal element As DependencyObject, ByVal value As Brush)
        If element Is Nothing Then
            Throw New ArgumentNullException("element")
        End If

        element.SetValue(DiagramTextBackgroundProperty, value)
    End Sub

    Public Shared ReadOnly DiagramTextBackgroundProperty As _
                           DependencyProperty = DependencyProperty.RegisterAttached("DiagramTextBackground",
                           GetType(Brush), GetType(Designer),
                           New PropertyMetadata(AddressOf DiagramTextBackgroundChanged))

    Shared Sub DiagramTextBackgroundChanged(Diagram As DependencyObject, e As DependencyPropertyChangedEventArgs)
        Dim Pnl = Helper.GetDiagramPanel(Diagram)
        If Pnl Is Nothing Then Return

        Pnl.DiagramTextBlock.Background = e.NewValue
        If Pnl.DrawOutlineMenuItem.IsChecked Then Pnl.DiagramObj.OutlineText()
    End Sub

#End Region

#Region "DiagramTextForeground"

    Public Shared Function GetDiagramTextForeground(ByVal element As DependencyObject) As Brush
        If element Is Nothing Then
            Throw New ArgumentNullException("element")
        End If

        Return element.GetValue(DiagramTextForegroundProperty)
    End Function

    Public Shared Sub SetDiagramTextForeground(ByVal element As DependencyObject, ByVal value As Brush)
        If element Is Nothing Then
            Throw New ArgumentNullException("element")
        End If

        element.SetValue(DiagramTextForegroundProperty, value)
    End Sub

    Public Shared ReadOnly DiagramTextForegroundProperty As _
                           DependencyProperty = DependencyProperty.RegisterAttached("DiagramTextForeground",
                           GetType(Brush), GetType(Designer),
                           New PropertyMetadata(Brushes.Black, AddressOf DiagramTextForegroundChanged))

    Shared Sub DiagramTextForegroundChanged(Diagram As DependencyObject, e As DependencyPropertyChangedEventArgs)
        Dim Pnl = Helper.GetDiagramPanel(Diagram)
        If Pnl Is Nothing Then Return

        Pnl.DiagramTextBlock.Foreground = e.NewValue
        If Pnl.DrawOutlineMenuItem.IsChecked Then Pnl.DiagramObj.OutlineText()
    End Sub

#End Region

#Region "DiagramTextOutlined"
    Public Shared Function GetDiagramTextOutlined(ByVal element As DependencyObject) As Boolean
        If element Is Nothing Then
            Throw New ArgumentNullException("element")
        End If

        Return element.GetValue(DiagramTextOutlinedProperty)
    End Function

    Public Shared Sub SetDiagramTextOutlined(ByVal element As DependencyObject, ByVal value As Boolean)
        If element Is Nothing Then
            Throw New ArgumentNullException("element")
        End If

        element.SetValue(DiagramTextOutlinedProperty, value)
    End Sub

    Public Shared ReadOnly DiagramTextOutlinedProperty As _
                           DependencyProperty = DependencyProperty.RegisterAttached("DiagramTextOutlined",
                           GetType(Boolean), GetType(Designer),
                           New PropertyMetadata(False, AddressOf DiagramTextOutlinedChanged))

    Shared Sub DiagramTextOutlinedChanged(Diagram As DependencyObject, e As DependencyPropertyChangedEventArgs)
        Dim Pnl = Helper.GetDiagramPanel(Diagram)
        If Pnl Is Nothing Then Return
        Pnl.DrawOutlineMenuItem.Tag = "DontUndo"
        Pnl.DrawOutlineMenuItem.IsChecked = e.NewValue
    End Sub

#End Region

#Region "DiagramTextOutlineThickness"
    Public Shared Function GetDiagramTextOutlineThickness(ByVal element As DependencyObject) As Double
        If element Is Nothing Then
            Throw New ArgumentNullException("element")
        End If

        Return element.GetValue(DiagramTextOutlineThicknessProperty)
    End Function

    Public Shared Sub SetDiagramTextOutlineThickness(ByVal element As DependencyObject, ByVal value As Double)
        If element Is Nothing Then
            Throw New ArgumentNullException("element")
        End If

        element.SetValue(DiagramTextOutlineThicknessProperty, value)
    End Sub

    Public Shared ReadOnly DiagramTextOutlineThicknessProperty As _
                           DependencyProperty = DependencyProperty.RegisterAttached("DiagramTextOutlineThickness",
                           GetType(Double), GetType(Designer),
                           New PropertyMetadata(1.0, AddressOf DiagramTextOutlineChanged))

#End Region

#Region "DiagramTextOutlineFill"


    Public Shared Function GetDiagramTextOutlineFill(ByVal element As DependencyObject) As Brush
        If element Is Nothing Then
            Throw New ArgumentNullException("element")
        End If

        Return element.GetValue(DiagramTextOutlineFillProperty)
    End Function

    Public Shared Sub SetDiagramTextOutlineFill(ByVal element As DependencyObject, ByVal value As Brush)
        If element Is Nothing Then
            Throw New ArgumentNullException("element")
        End If

        element.SetValue(DiagramTextOutlineFillProperty, value)
    End Sub

    Public Shared ReadOnly DiagramTextOutlineFillProperty As _
                           DependencyProperty = DependencyProperty.RegisterAttached("DiagramTextOutlineFill",
                           GetType(Brush), GetType(Designer),
                           New PropertyMetadata(AddressOf DiagramTextOutlineChanged))

    Shared Sub DiagramTextOutlineChanged(Diagram As DependencyObject, e As DependencyPropertyChangedEventArgs)
        Dim Pnl = Helper.GetDiagramPanel(Diagram)
        If Pnl Is Nothing Then Return
        If Pnl.DrawOutlineMenuItem.IsChecked Then Pnl.DiagramObj.OutlineText()

    End Sub

#End Region

#Region "DiagramTextApplyRotation"
    Public Shared Function GetDiagramTextApplyRotation(ByVal element As DependencyObject) As Boolean
        If element Is Nothing Then
            Throw New ArgumentNullException("element")
        End If

        Return element.GetValue(DiagramTextApplyRotationProperty)
    End Function

    Public Shared Sub SetDiagramTextApplyRotation(ByVal element As DependencyObject, ByVal value As Boolean)
        If element Is Nothing Then
            Throw New ArgumentNullException("element")
        End If

        element.SetValue(DiagramTextApplyRotationProperty, value)
    End Sub

    Public Shared ReadOnly DiagramTextApplyRotationProperty As _
                           DependencyProperty = DependencyProperty.RegisterAttached("DiagramTextApplyRotation",
                           GetType(Boolean), GetType(Designer),
                           New PropertyMetadata(True, AddressOf DiagramTextApplyRotationChanged))

    Shared Sub DiagramTextApplyRotationChanged(Diagram As DependencyObject, e As DependencyPropertyChangedEventArgs)
        Dim Pnl = Helper.GetDiagramPanel(Diagram)
        If Pnl Is Nothing Then Return
        Pnl.ApplyRotationMenuItem.IsChecked = e.NewValue
    End Sub

#End Region

#Region "HasChanges"

    Public Property HasChanges As Boolean
        Get
            Return GetValue(HasChangesProperty)
        End Get

        Set(ByVal value As Boolean)
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

        Set(ByVal value As Boolean)
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

        Set(ByVal value As Boolean)
            SetValue(CanRedoProperty, value)
        End Set
    End Property

    Public Property Text As String

    Public Shared ReadOnly CanRedoProperty As DependencyProperty =
                           DependencyProperty.Register("CanRedo",
                           GetType(Boolean), GetType(Designer),
                           New PropertyMetadata(False))

#End Region

#Region "GroupID"

    Public Shared Function GetGroupID(ByVal element As DependencyObject) As String
        If element Is Nothing Then
            Throw New ArgumentNullException("element")
        End If

        Return element.GetValue(GroupIDProperty)
    End Function

    Public Shared Sub SetGroupID(ByVal element As DependencyObject, ByVal value As String)
        If element Is Nothing Then
            Throw New ArgumentNullException("element")
        End If

        element.SetValue(GroupIDProperty, value)

    End Sub

    Public Shared ReadOnly GroupIDProperty As _
                           DependencyProperty = DependencyProperty.RegisterAttached("GroupID",
                           GetType(String), GetType(Designer),
                           New PropertyMetadata(AddressOf GroupIDChanged))

    Shared Sub GroupIDChanged(Diagram As DependencyObject, e As DependencyPropertyChangedEventArgs)
        Dim Pnl = Helper.GetDiagramPanel(Diagram)
        If Pnl Is Nothing Then Return
        Dim d As Date? = Nothing
        If e.NewValue <> "" Then d = Date.Parse(e.NewValue)
        DiagramGroup.Add(Pnl, d)
    End Sub

#End Region


#End Region




End Class
