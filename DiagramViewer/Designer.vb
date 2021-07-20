
Imports System.ComponentModel
Imports System.Windows.Markup
Imports System.Xml

Public Class Designer
    Inherits ListBox

    Friend ConnectionCanvas As Canvas
    Friend DesignerCanvas As Canvas
    Friend MaxZIndex As Integer = 0
    Friend SelectedBounds As Rect
    Friend GridPen As Pen

    Public Property FileName As String

    Public Sub New()
        Dim resourceLocater As Uri = New Uri("/DiagramViewer;component/Resources/designerdecorator.xaml", System.UriKind.Relative)
        Dim ResDec As ResourceDictionary = Application.LoadComponent(resourceLocater)
        Me.Resources.MergedDictionaries.Add(ResDec)
        Me.Style = FindResource("CanvasListBoxStyle")
        Me.ItemContainerStyle = FindResource("listBoxItemStyle")
    End Sub

    Private Shared NewPageOpened As Boolean

    Private Sub Designer_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        Dim Brdr As Border = VisualTreeHelper.GetChild(Me, 0)
        Dim G = CType(Brdr.Child, Grid)
        Me.ScrollViewer = G.Children(0)
        G = ScrollViewer.Content
        Me.DesignerCanvas = VisualTreeHelper.GetChild(G.Children(0), 0)
        DesignerGrid = G
        Me.ConnectionCanvas = G.Children(1)
        Me.DesignerCanvas.Width = Me.PageWidth
        Me.DesignerCanvas.Height = Me.PageHeight
        Me.ConnectionCanvas.Width = Me.PageWidth
        Me.ConnectionCanvas.Height = Me.PageHeight
        Me.DesignerCanvas.Background = SystemColors.ControlBrush

        If Not NewPageOpened Then
            NewPageOpened = True
        End If
        Me.Focus()

    End Sub

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
            Throw New Exception("Name can't start with a number.")
        End If

        For Each cnt In Me.Items
            If Me.GetControlName(cnt).ToLower() = newName Then
                Throw New Exception($"There is another control named '{newName}'!")
                Return False
            End If
        Next

        Try
            Dim fw = TryCast(control, FrameworkElement)
            If fw IsNot Nothing Then fw.Name = name
            Automation.AutomationProperties.SetName(control, name)
            Return True
        Catch ex As Exception
            Throw New Exception(ex.Message)
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

        Try
            Me.Name = name
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

    Private Function PageToXaml() As String

        Dim canvas As New Canvas
        canvas.Name = Me.Name
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

    Protected Overrides Sub OnItemsChanged(e As Specialized.NotifyCollectionChangedEventArgs)
        MyBase.OnItemsChanged(e)
        If e.Action = Specialized.NotifyCollectionChangedAction.Remove Then
            For Each Diagram In e.OldItems
                Dim Pnl = Helper.GetDiagramPanel(Diagram)
                Pnl.Dispose()
            Next
        End If
    End Sub


    Private Function GetNextName(name As String) As String
        Dim baseName As String
        Dim num As Integer
        For i = name.Length - 1 To 0 Step -1
            If Not IsNumeric(name(i)) Then
                If i = name.Length - 1 Then
                    baseName = name
                    num = 1
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
                If IsNumeric(n) Then
                    If CInt(n) >= num Then num = CInt(n) + 1
                End If
            End If
        Next
        Return baseName & num

    End Function


    Public Function Save() As Boolean
        If _FileName = "" Then
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
        If _FileName = "" Then
            saveName = Me.Name
        Else
            dlg.InitialDirectory = _FileName
            saveName = IO.Path.GetFileNameWithoutExtension(_FileName)
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
            Dim saveTo = If(tmpPath = "", _FileName, tmpPath)
            IO.File.WriteAllText(saveTo, xmal)
            If tmpPath = "" Then
                Me.HasChanges = False
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
            Return False
        End Try
        Return True
    End Function

    Public Sub Open()
        ' Configure open file dialog box 
        Dim dlg As New Microsoft.Win32.OpenFileDialog()
        dlg.DefaultExt = ".xaml" ' Default file extension
        dlg.Filter = "Diagram Pages|*.xaml"
        dlg.Title = "Open Diagram Design Page"
        If dlg.ShowDialog() = True Then
            Open(dlg.FileName)
            Me.HasChanges = False
        End If
    End Sub


    Public Sub Open(fileName As String)
        If Not IO.File.Exists(fileName) Then
            MsgBox($"File '{fileName}' doesn't exist!")
            Return
        End If

        Try
            Dim xaml = IO.File.ReadAllText(fileName)
            Me.XamlToPage(xaml)
            Me._FileName = IO.Path.GetFullPath(fileName)
        Catch ex As Exception
            MsgBox(ex.Message)
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


    Private Sub Designer_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles Me.MouseDown
        Me.SelectedIndex = -1
        Me.Focus()
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
            _Text = value
        Else
            SetControlText(Me.Items(controlIndex), value, True, True)
        End If
    End Sub

    Friend Sub SetControlText(control As UIElement, value As String, Optional trySetText As Boolean = True, Optional reportChanges As Boolean = False)
        If control Is Nothing Then control = Me.SelectedItem

        Dim txt = GetControlText(control)
        If txt = value Then Return

        If trySetText Then
            Try
                CObj(control).Text = value
                control.ClearValue(DiagramTextProperty)
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
        If GetDiagramTextOutlined(Diagram) Then Pnl.DiagramObj.OutlineText()
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

        If GetDiagramTextOutlined(Diagram) Then Pnl.DiagramObj.OutlineText()
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
        If GetDiagramTextOutlined(Diagram) Then Pnl.DiagramObj.OutlineText()
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
        If GetDiagramTextOutlined(Diagram) Then Pnl.DiagramObj.OutlineText()
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
                           GetType(Boolean), GetType(Designer))

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
        If GetDiagramTextOutlined(Diagram) Then Pnl.DiagramObj.OutlineText()

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
                           GetType(Boolean), GetType(Designer))

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
                           GetType(String), GetType(Designer))


#End Region


#End Region




End Class
