Imports System.Windows.Controls.Primitives
Imports System.Windows.Threading

Friend Class ColorList
    Inherits ListBox

    Dim WithEvents Pup As New Popup
    Dim WithEvents Lst As New ListBox
    Dim WithEvents Pkr As ColorPicker
    Dim Scv As ScrollViewer

    Public Property SelectedColor As Color
        Get
            Return GetValue(SelectedColorProperty)
        End Get

        Set(ByVal value As Color)
            SetValue(SelectedColorProperty, value)
        End Set
    End Property

    Public Shared ReadOnly SelectedColorProperty As DependencyProperty = _
                           DependencyProperty.Register("SelectedColor", _
                           GetType(Color), GetType(ColorList), _
                           New PropertyMetadata(Nothing))

    Private Sub ColorList_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded

        Pup.Child = Lst
        Me.Items.Clear()
        Me.Items.Add(New ColorLable(Colors.White))
        Me.Items.Add(New ColorLable(Colors.Cornsilk))
        Me.Items.Add(New ColorLable(Colors.Yellow))
        Me.Items.Add(New ColorLable(Colors.Gold))
        Me.Items.Add(New ColorLable(Colors.BurlyWood))
        Me.Items.Add(New ColorLable(Colors.Khaki))
        Me.Items.Add(New ColorLable(Colors.Orange))
        Me.Items.Add(New ColorLable(Colors.Violet))
        Me.Items.Add(New ColorLable(Colors.Purple))
        Me.Items.Add(New ColorLable(Colors.MediumVioletRed))
        Me.Items.Add(New ColorLable(Colors.IndianRed))
        Me.Items.Add(New ColorLable(Colors.Red))
        Me.Items.Add(New ColorLable(Colors.Green))
        Me.Items.Add(New ColorLable(Colors.Cyan))
        Me.Items.Add(New ColorLable(Colors.Azure))
        Me.Items.Add(New ColorLable(Colors.LightBlue))
        Me.Items.Add(New ColorLable(Colors.LightSkyBlue))
        Me.Items.Add(New ColorLable(Colors.DeepSkyBlue))
        Me.Items.Add(New ColorLable(Colors.CornflowerBlue))
        Me.Items.Add(New ColorLable(Colors.Blue))

        Scv = GetParent(Me, GetType(ScrollViewer))
        Pkr = GetParent(Scv, GetType(ColorPicker))
    End Sub

    Function GetParent(Child As UIElement, ParentType As Type)
        Dim P = Child
        Do
            P = VisualTreeHelper.GetParent(P)
            If P Is Nothing OrElse P.GetType Is ParentType Then Return P
        Loop
    End Function

    Private Sub ColorList_MouseLeftButtonUp(sender As Object, e As MouseButtonEventArgs) Handles Me.MouseLeftButtonUp
        ShowPopUp()
    End Sub

    Private Sub ColorList_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles Me.SelectionChanged
        Dim ClrLb As ColorLable = Me.SelectedItem

        For Each Item As ColorLable In Me.Items
            Item.IsSelected = (Item Is ClrLb)
        Next

        If ClrLb Is Nothing Then Return

        ItIsMe = True
        Me.SelectedColor = ClrLb.Color
    End Sub

    Private Sub Lst_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles Lst.SelectionChanged
        If Lst.SelectedIndex = -1 Then Return
        Dim ClrLb = CType(Lst.SelectedItem, ColorLable)
        For Each Item As ColorLable In Lst.Items
            Item.IsSelected = (Item Is ClrLb)
        Next

        ItIsMe = True
        Me.SelectedColor = ClrLb.Color        
    End Sub

    Dim WasOpen As Boolean = False
    Dim ColorPicker As ColorPicker

    Private Sub ShowPopUp()         
        If Not Pkr.ShowPopUpColorDegrees Then Return

        Dim ClrLb As ColorLable = Me.SelectedItem
        If ClrLb Is Nothing Then Return

        If Pup.PlacementTarget Is ClrLb AndAlso WasOpen Then
            WasOpen = False
            Return
        End If

        Pup.IsOpen = False
        Pup.StaysOpen = False

        Lst.Items.Clear()


        Dim C = ClrLb.Color
        Dim H, S, B As Double
        ColorHelper.HSBFromColor(C, H, S, B)
        For i = 12 To 0 Step -1
            If i < 4 Then Exit For
            Lst.Items.Add(New ColorLable(H, S, i / 12))
        Next

        Pup.PlacementTarget = ClrLb
        Pup.HorizontalOffset = -6
        Pup.IsOpen = True
        Lst.SelectedIndex = (1 - B) * 12
        Dim Item As ListBoxItem = GetParent(Lst.SelectedItem, GetType(ListBoxItem))
        If Item IsNot Nothing Then Item.Focus()
    End Sub

    Friend Sub ComputeSize()
        If Me.Items.Count = 0 Then Return
        Dim PART_WrapPanel As WrapPanel = GetParent(Me.Items(0), GetType(WrapPanel))
        Dim ScWidth = If(Scv.ComputedVerticalScrollBarVisibility = Windows.Visibility.Visible, SystemParameters.ScrollWidth, 0)
        Dim NewWidth = PART_WrapPanel.ActualWidth / 10 - 6.4

        For Each Item As ColorLable In Me.Items
            Item.Width = NewWidth
            Item.Height = NewWidth
        Next

    End Sub

    Private Sub Pup_Closed(sender As Object, e As EventArgs) Handles Pup.Closed
        WasOpen = (Keyboard.FocusedElement() Is GetParent(Pup.PlacementTarget, GetType(ListBoxItem)))
    End Sub

    Dim ItIsMe As Boolean = False
    Private Sub Pkr_ColorChanged(sender As Object, e As ColorChangedEventArgs) Handles Pkr.ColorChanged
        If ItIsMe Then
            ItIsMe = False
            Return
        End If

        For Each Item As ColorLable In Me.Items
            Item.IsSelected = False
        Next
    End Sub

End Class
