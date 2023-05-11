Imports System.Windows.Markup

Friend Class TileBrushPicker

    Private Property HatchBrush As DrawingBrush
        Get
            Return GetValue(HatchBrushProperty)
        End Get

        Set(ByVal value As DrawingBrush)
            SetValue(HatchBrushProperty, value)
        End Set
    End Property

    Public Shared ReadOnly HatchBrushProperty As DependencyProperty = _
                           DependencyProperty.Register("HatchBrush", _
                           GetType(DrawingBrush), GetType(TileBrushPicker))

    Public Property Brush As Brush
        Get
            Return GetValue(BrushProperty)
        End Get

        Set(ByVal value As Brush)
            SetValue(BrushProperty, value)
        End Set
    End Property

    Public Shared ReadOnly BrushProperty As DependencyProperty = _
                           DependencyProperty.Register("Brush", _
                           GetType(Brush), GetType(TileBrushPicker))

    Dim KeepHatchBrush As Boolean = False

    Private Sub LstTile_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles LstTile.SelectionChanged
        If LstTile.SelectedIndex = -1 Then
            StkPnlTileOptions.IsEnabled = False
            Return
        End If
        StkPnlTileOptions.IsEnabled = True

        If Not KeepHatchBrush Then
            Dim Hb As DrawingBrush = LstTile.SelectedItem
            Dim P As New HatchParams(HatchBrushes.GetHatchStyle(Hb), PkrBackground.Brush, PkrForeground.Brush, UdPenThickness.Value, HatchTransPkr.Transform, ShapeTransPkr.Transform)
            HatchBrush = New DrawingBrush
            HatchBrushes.SetHatchBrushParam(HatchBrush, P)

            HatchTransPkr.Transform = Nothing
            HatchTransPkr.TargetWidth = HatchBrushes.GetHatchWidth(HatchBrush)
            HatchTransPkr.TargetHeight = HatchBrushes.GetHatchHeight(HatchBrush)

            ShapeTransPkr.Transform = Nothing
            ShapeTransPkr.TargetWidth = HatchBrushes.GetHatchWidth(HatchBrush)
            ShapeTransPkr.TargetHeight = HatchBrushes.GetHatchHeight(HatchBrush)

            Me.Brush = HatchBrush
        End If

        UpdateControls()
    End Sub

    Private Sub PkrBackground_BrushChanged(OldBrush As Brush, NewBrush As Brush) Handles PkrBackground.BrushChanged
        If HatchBrush IsNot Nothing Then HatchBrushes.SetBackground(HatchBrush, NewBrush)
        Me.Brush = HatchBrush
    End Sub

    Private Sub PkrForeground_BrushChanged(OldBrush As Brush, NewBrush As Brush) Handles PkrForeground.BrushChanged
        If HatchBrush IsNot Nothing Then HatchBrushes.SetForeground(HatchBrush, NewBrush)
        Me.Brush = HatchBrush
    End Sub

    Public Property PreviewBox As ContentControl
        Get
            Return GetValue(PreviewBoxProperty)
        End Get

        Set(ByVal value As ContentControl)
            SetValue(PreviewBoxProperty, value)
        End Set
    End Property

    Public Shared ReadOnly PreviewBoxProperty As DependencyProperty = _
                           DependencyProperty.Register("PreviewBox", _
                           GetType(Control), GetType(TileBrushPicker), _
                           New PropertyMetadata(Nothing, AddressOf OnPreviewBoxChanged))

    Shared Sub OnPreviewBoxChanged(TPkr As TileBrushPicker, e As DependencyPropertyChangedEventArgs)
        Dim Pv As ContentControl = e.NewValue
        TPkr.PkrBackground.PreviewBox = Pv
        TPkr.PkrForeground.PreviewBox = Pv
        Dim OldPv As ContentControl = e.OldValue
        If OldPv IsNot Nothing Then OldPv.Content = Nothing
    End Sub

    Sub SetParams()
        If LstTile.ItemsSource Is Nothing Then
            LstTile.ItemsSource = HatchBrushes.GetAllHatchBrushes
        ElseIf Not TypeOf Me.Brush Is DrawingBrush Then
            LstTile.SelectedIndex = -1
        End If

        StkPnlTileOptions.IsEnabled = (LstTile.SelectedIndex <> -1)
        If Me.Brush Is Nothing OrElse Not TypeOf Me.Brush Is DrawingBrush Then Return

        KeepHatchBrush = True
        HatchBrush = Me.Brush

        Dim B = HatchBrushes.GetBackground(HatchBrush)
        PkrBackground.Brush = B ' If(B Is Nothing, Nothing, B.CloneCurrentValue)

        B = HatchBrushes.GetForeground(HatchBrush)
        PkrForeground.Brush = B ' If(B Is Nothing, Nothing, B.CloneCurrentValue)

        UdPenThickness.Value = HatchBrushes.GetPenThickness(HatchBrush)
        HatchTransPkr.Transform = HatchBrushes.GetHatchTransform(HatchBrush)
        ShapeTransPkr.Transform = HatchBrushes.GetShapeTransform(HatchBrush)

        LstTile.SelectedIndex = HatchBrushes.GetHatchStyle(HatchBrush)
        LstTile.ScrollIntoView(LstTile.SelectedItem)
        KeepHatchBrush = False
    End Sub

    Private Sub UdViewBoxHeight_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Object)) Handles UdViewBoxHeight.ValueChanged
        HatchBrush.ViewboxUnits = BrushMappingMode.Absolute
        HatchBrush.Viewbox = New Rect(
            HatchBrush.Viewbox.X,
            HatchBrush.Viewbox.Y,
            HatchBrush.Viewbox.Width,
            If(UdViewBoxHeight.Value, Double.NaN)
        )
    End Sub

    Private Sub UdViewBoxWidth_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Object)) Handles UdViewBoxWidth.ValueChanged
        HatchBrush.ViewboxUnits = BrushMappingMode.Absolute
        HatchBrush.Viewbox = New Rect(
            HatchBrush.Viewbox.X,
            HatchBrush.Viewbox.Y,
            If(UdViewBoxWidth.Value, Double.NaN),
            HatchBrush.Viewbox.Height
       )
    End Sub

    Private Sub UdViewBoxX_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Object)) Handles UdViewBoxX.ValueChanged
        HatchBrush.ViewboxUnits = BrushMappingMode.Absolute
        HatchBrush.Viewbox = New Rect(
            If(UdViewBoxX.Value, 0),
            HatchBrush.Viewbox.Y,
            HatchBrush.Viewbox.Width,
            HatchBrush.Viewbox.Height
       )
    End Sub

    Private Sub UdViewBoxY_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Object)) Handles UdViewBoxY.ValueChanged
        HatchBrush.ViewboxUnits = BrushMappingMode.Absolute
        HatchBrush.Viewbox = New Rect(
            HatchBrush.Viewbox.X,
            If(UdViewBoxY.Value, 0),
            HatchBrush.Viewbox.Width,
            HatchBrush.Viewbox.Height
       )
    End Sub

    Private Sub UdViewportHeight_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Object)) Handles UdViewPortHeight.ValueChanged
        HatchBrush.ViewportUnits = BrushMappingMode.Absolute
        HatchBrush.Viewport = New Rect(
            HatchBrush.Viewport.X,
            HatchBrush.Viewport.Y,
            HatchBrush.Viewport.Width,
            If(UdViewPortHeight.Value, Double.NaN)
        )
    End Sub

    Private Sub UdViewportWidth_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Object)) Handles UdViewPortWidth.ValueChanged
        HatchBrush.ViewportUnits = BrushMappingMode.Absolute
        HatchBrush.Viewport = New Rect(
            HatchBrush.Viewport.X,
            HatchBrush.Viewport.Y,
            If(UdViewPortWidth.Value, Double.NaN),
            HatchBrush.Viewport.Height
        )
    End Sub

    Private Sub UdViewportX_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Object)) Handles UdViewPortX.ValueChanged
        HatchBrush.ViewportUnits = BrushMappingMode.Absolute
        HatchBrush.Viewport = New Rect(
            If(UdViewPortX.Value, 0),
            HatchBrush.Viewport.Y,
            HatchBrush.Viewport.Width,
            HatchBrush.Viewport.Height
        )
    End Sub

    Private Sub UdViewportY_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Object)) Handles UdViewPortY.ValueChanged
        HatchBrush.ViewportUnits = BrushMappingMode.Absolute
        HatchBrush.Viewport = New Rect(
            HatchBrush.Viewport.X,
            If(UdViewPortY.Value, 0),
            HatchBrush.Viewport.Width,
            HatchBrush.Viewport.Height
        )
    End Sub

    Private Sub UpdateControls()
        UdViewBoxX.Value = HatchBrush.Viewbox.X
        UdViewBoxY.Value = HatchBrush.Viewbox.Y
        UdViewBoxWidth.Value = HatchBrush.Viewbox.Width
        UdViewBoxHeight.Value = HatchBrush.Viewbox.Height

        UdViewPortX.Value = HatchBrush.Viewport.X
        UdViewPortY.Value = HatchBrush.Viewport.Y
        UdViewPortWidth.Value = HatchBrush.Viewport.Width
        UdViewPortHeight.Value = HatchBrush.Viewport.Height

        cmbStretch.SelectedIndex = HatchBrush.Stretch
        cmbTileMode.SelectedIndex = HatchBrush.TileMode
        cmbAlignX.SelectedIndex = HatchBrush.AlignmentX
        cmbAlignY.SelectedIndex = HatchBrush.AlignmentY
    End Sub

    Private Sub TileBrushPicker_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        cmbStretch.ItemsSource = [Enum].GetNames(GetType(Stretch))
        cmbTileMode.ItemsSource = [Enum].GetNames(GetType(TileMode))
        cmbAlignX.ItemsSource = [Enum].GetNames(GetType(AlignmentX))
        cmbAlignY.ItemsSource = [Enum].GetNames(GetType(AlignmentY))
    End Sub

    Private Sub LstTile_SizeChanged(sender As Object, e As SizeChangedEventArgs) Handles LstTile.SizeChanged
        Dim W = LstTile.ActualWidth / 5 - 17
        For Each Brush As Brush In LstTile.Items
            Dim ListBoxItem As ListBoxItem = LstTile.ItemContainerGenerator.ContainerFromItem(Brush)
            Dim Bd = VisualTreeHelper.GetChild(ListBoxItem, 0)
            Dim Cp = VisualTreeHelper.GetChild(Bd, 0)
            Dim Brdr As Border = VisualTreeHelper.GetChild(Cp, 0)
            Brdr.Width = W
            Brdr.Height = W
        Next
    End Sub

    Private Sub UdLineThickness_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Object)) Handles UdPenThickness.ValueChanged
        If HatchBrush IsNot Nothing Then HatchBrushes.SetPenThickness(HatchBrush, UdPenThickness.Value)
        Me.Brush = HatchBrush
    End Sub

    Private Sub cmbStretch_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles cmbStretch.SelectionChanged
        HatchBrush.Stretch = cmbStretch.SelectedIndex
    End Sub

    Private Sub cmbTileMode_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles cmbTileMode.SelectionChanged
        HatchBrush.TileMode = cmbTileMode.SelectedIndex
    End Sub

    Private Sub cmbAlignX_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles cmbAlignX.SelectionChanged
        HatchBrush.AlignmentX = cmbAlignX.SelectedIndex
    End Sub

    Private Sub cmbAlignY_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles cmbAlignY.SelectionChanged
        HatchBrush.AlignmentY = cmbAlignY.SelectedIndex
    End Sub

    Private Sub HatchTransPkr_TransformChanged(TransBox As TransformBox, OldTransform As Transform, NewTransform As Transform) Handles HatchTransPkr.TransformChanged
        HatchBrushes.SetHatchTransform(Me.HatchBrush, NewTransform)
    End Sub

    Private Sub ShapeTransPkr_TransformChanged(TransBox As TransformBox, OldTransform As Transform, NewTransform As Transform) Handles ShapeTransPkr.TransformChanged
        HatchBrushes.SetShapeTransform(Me.HatchBrush, NewTransform)
    End Sub

    Private Sub Pkr_Opend() Handles PkrForeground.Opening, PkrBackground.Opening
        WndColor.KeepLstBrushesFilter = True
    End Sub

    Private Sub Pkr_Closed(Canceled As Boolean) Handles PkrForeground.Closed, PkrBackground.Closed
        PreviewBox.Background = Me.Brush
        WndColor.KeepLstBrushesFilter = False
    End Sub

End Class
