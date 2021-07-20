Imports System.ComponentModel

Public Class DiagramPanel
    Inherits UserControl

    Friend ConnectorsGrid As Grid
    Friend DiagramTextBlock As TextBlock
    Friend MyDesigner As Designer
    Friend DesignerItem As ListBoxItem
    Friend Diagram As FrameworkElement
    Dim Scv As ScrollViewer
    Friend DiagramObj As DiagramObject
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

        Diagram = TryCast(CType(Me.Content, ContentPresenter).Content, FrameworkElement)
        If Diagram Is Nothing Then Return

        Canvas.SetLeft(DesignerItem, Designer.GetLeft(Diagram))
        Canvas.SetRight(DesignerItem, Designer.GetRight(Diagram))
        Canvas.SetTop(DesignerItem, Designer.GetTop(Diagram))
        Canvas.SetBottom(DesignerItem, Designer.GetBottom(Diagram))

        Me.Width = Helper.FixToMm(Designer.GetFrameWidth(Diagram)) + Diagram.Margin.Left + Diagram.Margin.Right
        Me.Height = Helper.FixToMm(Designer.GetFrameHeight(Diagram)) + Diagram.Margin.Top + Diagram.Margin.Bottom

        Dim Angle = Designer.GetRotationAngle(Diagram)
        Helper.Rotate(DesignerItem, Angle)

        DiagramTextBlock = Me.Template.FindName("PART_Text", Me)
        ConnectorsGrid = DiagramTextBlock.Parent
        Designer.SetDiagramTextBlock(Diagram, DiagramTextBlock)

        AddHandler DesignerItem.GotFocus, AddressOf DesignerItem_GotFocus
        AddHandler Diagram.PreviewKeyDown, AddressOf DesignerItem_PreviewKeyDown
    End Sub

    Private Sub DiagramPanel_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        DiagramObj = DiagramObject.CreateDiagramObject(Diagram)

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

        Dim sk = TryCast(Diagram.LayoutTransform, SkewTransform)
        If sk Is Nothing OrElse sk.AngleY = 0 Then Diagram.LayoutTransform = New SkewTransform(0, 0.000000000001, 0, 0.000000000001)

        TopDpd = DependencyPropertyDescriptor.FromProperty(Canvas.TopProperty, GetType(ListBoxItem))
        TopDpd.AddValueChanged(DesignerItem, AddressOf DesignerItem_TopChanged)

        LeftDpd = DependencyPropertyDescriptor.FromProperty(Canvas.LeftProperty, GetType(ListBoxItem))
        LeftDpd.AddValueChanged(DesignerItem, AddressOf DesignerItem_LeftChanged)

        RenderTransformDpd = DependencyPropertyDescriptor.FromProperty(ListBoxItem.RenderTransformProperty, GetType(ListBoxItem))
        RenderTransformDpd.AddValueChanged(DesignerItem, AddressOf DesignerItem_RenderTransformChanged)

        TextDpd = DependencyPropertyDescriptor.FromProperty(TextBlock.TextProperty, GetType(TextBlock))

    End Sub

    Sub Dispose()
        DiagramGroup.RemovePanelOnly(Me)

        RemoveHandler DesignerItem.PreviewKeyDown, AddressOf DesignerItem_PreviewKeyDown
        RemoveHandler DesignerItem.GotFocus, AddressOf DesignerItem_GotFocus

        TopDpd.RemoveValueChanged(DesignerItem, AddressOf DesignerItem_TopChanged)
        LeftDpd.RemoveValueChanged(DesignerItem, AddressOf DesignerItem_LeftChanged)
        RenderTransformDpd.RemoveValueChanged(DesignerItem, AddressOf DesignerItem_RenderTransformChanged)

        Dim Cp As ContentPresenter = VisualTreeHelper.GetParent(Diagram)
        Cp.Content = Nothing
        ConnectorsGrid = Nothing
        DiagramTextBlock = Nothing
        DesignerItem = Nothing
        Diagram = Nothing
        Scv = Nothing
        DiagramObj = Nothing
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


#Region "DesignerItem"

    Private Sub DesignerItem_GotFocus(sender As Object, e As RoutedEventArgs)
        MyDesigner.MaxZIndex += 1
        Canvas.SetZIndex(DesignerItem, MyDesigner.MaxZIndex)
        MyDesigner.ScrollIntoView(Diagram)
    End Sub

    Private Sub DesignerItem_PreviewKeyDown(sender As Object, e As KeyEventArgs)
        If DesignerItem Is Nothing Then Return

        Select Case e.Key
            Case Key.Tab
                e.Handled = True

                Dim index = MyDesigner.Items.IndexOf(Diagram)
                If Keyboard.Modifiers = ModifierKeys.Shift Then
                    For i = index - 1 To 0 Step -1
                        Dim item As UIElement = MyDesigner.Items(i)
                        If item.Focusable Then
                            item.Focus()
                            Return
                        End If
                    Next

                    For i = MyDesigner.Items.Count - 1 To index Step -1
                        Dim item As UIElement = MyDesigner.Items(i)
                        If item.Focusable Then
                            item.Focus()
                            Return
                        End If
                    Next
                Else
                    For i = index + 1 To MyDesigner.Items.Count - 1
                        Dim item As UIElement = MyDesigner.Items(i)
                        If item.Focusable Then
                            item.Focus()
                            Return
                        End If
                    Next

                    For i = 0 To index
                        Dim item As UIElement = MyDesigner.Items(i)
                        If item.Focusable Then
                            item.Focus()
                            Return
                        End If
                    Next
                End If

        End Select
    End Sub

    Dim TopDpd As DependencyPropertyDescriptor
    Dim LeftDpd As DependencyPropertyDescriptor
    Dim RenderTransformDpd As DependencyPropertyDescriptor
    Dim TextDpd As DependencyPropertyDescriptor
    Friend FocusRectangle As New Border

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
            If Not Designer.GetDiagramTextApplyRotation(Diagram) Then DiagramTextBlock.RenderTransform = DesignerItem.RenderTransform.Inverse
        End If
    End Sub

#End Region

End Class
