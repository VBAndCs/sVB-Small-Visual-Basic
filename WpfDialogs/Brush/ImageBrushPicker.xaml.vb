Imports System.Windows.Markup

Friend Class ImageBrushPicker

    Private Property ImageBrush As ImageBrush
        Get
            Return GetValue(ImageBrushProperty)
        End Get

        Set(ByVal value As ImageBrush)
            SetValue(ImageBrushProperty, value)
        End Set
    End Property

    Public Shared ReadOnly ImageBrushProperty As DependencyProperty = _
                           DependencyProperty.Register("ImageBrush", _
                           GetType(ImageBrush), GetType(ImageBrushPicker))

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
                           GetType(Brush), GetType(ImageBrushPicker))

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
                           GetType(Control), GetType(ImageBrushPicker))

    Sub SetParams()
        If Me.Brush Is Nothing OrElse Not TypeOf Me.Brush Is ImageBrush Then
            txtFileName.Text = ""
        Else
            ImageBrush = Me.Brush
            txtFileName.Text = Me.ImageBrush.ImageSource.ToString            
            UpdateControls()
        End If
    End Sub

    Private Sub UdViewBoxHeight_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Object)) Handles UdViewBoxHeight.ValueChanged
        ImageBrush.Viewbox = New Rect(ImageBrush.Viewbox.X, ImageBrush.Viewbox.Y, ImageBrush.Viewbox.Width, UdViewBoxHeight.Value)
    End Sub

    Private Sub UdViewBoxWidth_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Object)) Handles UdViewBoxWidth.ValueChanged
        ImageBrush.Viewbox = New Rect(ImageBrush.Viewbox.X, ImageBrush.Viewbox.Y, UdViewBoxWidth.Value, ImageBrush.Viewbox.Height)
    End Sub

    Private Sub UdViewBoxX_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Object)) Handles UdViewBoxX.ValueChanged
        ImageBrush.Viewbox = New Rect(UdViewBoxX.Value, ImageBrush.Viewbox.Y, ImageBrush.Viewbox.Width, ImageBrush.Viewbox.Height)
    End Sub

    Private Sub UdViewBoxY_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Object)) Handles UdViewBoxY.ValueChanged
        ImageBrush.Viewbox = New Rect(ImageBrush.Viewbox.X, UdViewBoxY.Value, ImageBrush.Viewbox.Width, ImageBrush.Viewbox.Height)
    End Sub


    Private Sub UdViewportHeight_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Object)) Handles UdViewPortHeight.ValueChanged
        ImageBrush.Viewport = New Rect(
            ImageBrush.Viewport.X,
            ImageBrush.Viewport.Y,
            ImageBrush.Viewport.Width,
            If(UdViewPortHeight.Value, Double.NaN)
        )
    End Sub

    Private Sub UdViewportWidth_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Object)) Handles UdViewPortWidth.ValueChanged
        ImageBrush.Viewport = New Rect(
            ImageBrush.Viewport.X,
            ImageBrush.Viewport.Y,
           If(UdViewPortWidth.Value, Double.NaN),
            ImageBrush.Viewport.Height
        )
    End Sub

    Private Sub UdViewportX_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Object)) Handles UdViewPortX.ValueChanged
        ImageBrush.Viewport = New Rect(
            If(UdViewPortX.Value, 0),
            ImageBrush.Viewport.Y,
            ImageBrush.Viewport.Width,
            ImageBrush.Viewport.Height
        )
    End Sub

    Private Sub UdViewportY_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Object)) Handles UdViewPortY.ValueChanged
        ImageBrush.Viewport = New Rect(
            ImageBrush.Viewport.X,
            If(UdViewPortY.Value, 0),
            ImageBrush.Viewport.Width,
            ImageBrush.Viewport.Height
        )
    End Sub

    Private Sub UpdateControls()
        UdViewBoxX.Value = ImageBrush.Viewbox.X
        UdViewBoxY.Value = ImageBrush.Viewbox.Y
        UdViewBoxWidth.Value = ImageBrush.Viewbox.Width
        UdViewBoxHeight.Value = ImageBrush.Viewbox.Height

        UdViewPortX.Value = ImageBrush.Viewport.X
        UdViewPortY.Value = ImageBrush.Viewport.Y
        UdViewPortWidth.Value = ImageBrush.Viewport.Width
        UdViewPortHeight.Value = ImageBrush.Viewport.Height

        cmbStretch.SelectedIndex = ImageBrush.Stretch
        cmbTileMode.SelectedIndex = ImageBrush.TileMode
        cmbAlignX.SelectedIndex = ImageBrush.AlignmentX
        cmbAlignY.SelectedIndex = ImageBrush.AlignmentY
    End Sub

    Private Sub TileBrushPicker_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        cmbStretch.ItemsSource = [Enum].GetNames(GetType(Stretch))
        cmbTileMode.ItemsSource = [Enum].GetNames(GetType(TileMode))
        cmbAlignX.ItemsSource = [Enum].GetNames(GetType(AlignmentX))
        cmbAlignY.ItemsSource = [Enum].GetNames(GetType(AlignmentY))
    End Sub

    Private Sub cmbStretch_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles cmbStretch.SelectionChanged
        ImageBrush.Stretch = cmbStretch.SelectedIndex
    End Sub

    Private Sub cmbTileMode_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles cmbTileMode.SelectionChanged
        ImageBrush.TileMode = cmbTileMode.SelectedIndex
    End Sub

    Private Sub cmbAlignX_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles cmbAlignX.SelectionChanged
        ImageBrush.AlignmentX = cmbAlignX.SelectedIndex
    End Sub

    Private Sub cmbAlignY_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles cmbAlignY.SelectionChanged
        ImageBrush.AlignmentY = cmbAlignY.SelectedIndex
    End Sub

    Dim TileImage As BitmapImage

    Private Sub btnOpenFile_Click(sender As Object, e As RoutedEventArgs)
        ' Configure open file dialog box 
        Dim dlg As New Microsoft.Win32.OpenFileDialog()
        dlg.FileName = txtFileName.Text
        dlg.DefaultExt = ".bmp" ' Default file extension
        dlg.Filter = "BMP Images|*.bmp|" &
                          "JPEG Images|*.jpg|" &
                          "GIF Images|*.gif|" &
                          "PNG Images|*.png|" &
                          "Icons|*.ico|" &
                          "TIFF Images|*.tiff|" &
                          "WDP Images|*.wdp|" &
                          "All Image Types|*.bmp;*.jpg;*.gif;*.png;*.tiff;*.ico;*.wdp"

        dlg.FilterIndex = 8

        ' Show open file dialog box 
        Dim result? As Boolean = dlg.ShowDialog()

        ' Process open file dialog box results 
        If result = True Then
            txtFileName.Text = dlg.FileName
            Me.ImageBrush = CreateImageBrush(dlg.FileName)
            Me.Brush = Me.ImageBrush
            UpdateControls()
        End If
    End Sub

    Private Sub txtFileName_TextChanged(sender As Object, e As TextChangedEventArgs)
        StkPnlTileOptions.IsEnabled = (txtFileName.Text <> "")
    End Sub

    Friend Shared Function CreateImageBrush(FileName As String) As ImageBrush
        Dim Uri As New Uri(FileName, UriKind.Absolute)
        Return CreateImageBrush(Uri)
    End Function

    Friend Shared Function CreateImageBrush(Uri As Uri) As ImageBrush
        Dim ImageBrush As New ImageBrush
        Dim TileImage = New BitmapImage(Uri)
        ImageBrush.ImageSource = TileImage
        ImageBrush.TileMode = TileMode.Tile
        Return ImageBrush
    End Function

End Class
