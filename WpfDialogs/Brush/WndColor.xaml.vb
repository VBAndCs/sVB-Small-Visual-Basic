Imports System.Windows.Markup
Imports System.Xml
Imports System.Text
Imports System.ComponentModel


Friend Class WndColor

    Friend Shared KeepLstBrushesFilter As Boolean = False

    Public Property Brush() As Brush
        Get
            Return lblBrushPreview.Background
        End Get

        Set(ByVal value As Brush)
            If TypeOf value Is DrawingBrush Then
                TilePicker.Brush = value
                TabControl1.SelectedIndex = 1
            ElseIf TypeOf value Is ImageBrush Then
                ImagePicker.Brush = value
                TabControl1.SelectedIndex = 2
            Else
                Pkr.Brush = value
                TabControl1.SelectedIndex = 0
            End If
            GetBrushXaml()
        End Set
    End Property

    Private Sub BtnOk_Click(sender As Object, e As RoutedEventArgs)
        Me.DialogResult = True
        AddToRecent()
    End Sub

    Private Sub BtnCancel_Click(sender As Object, e As RoutedEventArgs)
        Me.DialogResult = False
    End Sub

    Private Sub TabControl1_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles TabControl1.SelectionChanged
        If e.OriginalSource IsNot TabControl1 Then Return

        Select Case TabControl1.SelectedIndex
            Case 1
                TilePicker.Brush = Me.Brush
                TilePicker.SetParams()
                lstBrushes.Items.Filter = Function(B As Brush) TypeOf B Is DrawingBrush
            Case 2
                ImagePicker.Brush = Me.Brush
                ImagePicker.SetParams()
                lstBrushes.Items.Filter = Function(B As Brush) TypeOf B Is ImageBrush
            Case Else
                If Pkr.Brush Is Nothing Then
                    lstBrushes.Items.Filter = Nothing
                Else
                    Dim T = Pkr.Brush.GetType
                    lstBrushes.Items.Filter = Function(B As Brush) B.GetType Is T
                End If

        End Select

    End Sub

    Private Sub GetBrushXaml()
        If Me.Brush Is Nothing Then
            lstBrushes.Items.Filter = Nothing
            txtXaml.Text = ""
            Return
        End If

        If Not KeepLstBrushesFilter Then
            Dim T = Me.Brush.GetType
            lstBrushes.Items.Filter = Function(B As Brush) B.GetType Is T
        End If

        Dim settings As New XmlWriterSettings()
        settings.Indent = True
        settings.OmitXmlDeclaration = True
        settings.NewLineOnAttributes = True

        Dim sb As New StringBuilder()
        Dim writer As XmlWriter = XmlWriter.Create(sb, settings)
        XamlWriter.Save(HatchBrushes.CloneDrawingBrush(Me.Brush), writer)
        sb.Replace(" xmlns=" & Chr(34) & "http://schemas.microsoft.com/winfx/2006/xaml/presentation" & Chr(34), "")
        sb.Replace(" xmlns:x=" & Chr(34) & "http://schemas.microsoft.com/winfx/2006/xaml" & Chr(34), "")
        txtXaml.Text = sb.ToString
    End Sub

    Private Sub WndColor_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        Dim dpd = DependencyPropertyDescriptor.FromProperty(Label.BackgroundProperty, GetType(Label))
        If (dpd IsNot Nothing) Then
            dpd.AddValueChanged(lblBrushPreview, AddressOf GetBrushXaml)
        End If

        Try
            Dim brushesStream = IO.File.OpenRead(RecentBrushedPath)
            Dim reader As New XamlReader()
            Dim Lst = CType(reader.LoadAsync(brushesStream), ListBox)
            brushesStream.Close()

            For i = 0 To Math.Min(Lst.Items.Count - 1, 150)
                Dim ImgBrush = TryCast(Lst.Items(i), ImageBrush)
                If ImgBrush IsNot Nothing Then
                    Dim x As String = ImageBrushes.GetImageFileName(ImgBrush)
                    If x <> "" Then
                        Dim uri As New Uri(x)
                        If IO.File.Exists(uri.AbsolutePath) Then
                            ImgBrush.ImageSource = New BitmapImage(uri)
                            lstBrushes.Items.Add(ImgBrush)
                        End If
                    End If
                Else
                    lstBrushes.Items.Add(Lst.Items(i))
                End If
            Next
            Dim T = Me.Brush?.GetType
            lstBrushes.Items.Filter = Function(B As Brush) B.GetType Is T
        Catch ex As Exception

        End Try
    End Sub

    Private Shared RecentBrushedPath As String = IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "RecentBrushes.xaml")

    Private Sub AddToRecent()
        If Me.Brush Is Nothing Then Return

        For Each B As Brush In lstBrushes.Items
            If BrushHelper.BrushesAreEqual(B, Me.Brush) Then
                lstBrushes.Items.Remove(B)
                Exit For
            End If
        Next

        lstBrushes.Items.Filter = Nothing
        lstBrushes.Items.Insert(0, HatchBrushes.Clone(Me.Brush))

        Dim Lst As New ListBox
        For Each B As Brush In lstBrushes.Items
            Dim ImgBrush = TryCast(B, ImageBrush)
            If ImgBrush IsNot Nothing Then
                ImageBrushes.SetImageFileName(ImgBrush, ImgBrush.ImageSource.ToString)
                ImgBrush.ImageSource = Nothing
                Lst.Items.Add(ImgBrush)
            Else
                Lst.Items.Add(B)
            End If
        Next

        Dim Xaml = XamlWriter.Save(Lst)
        Dim Sw As New IO.StreamWriter(RecentBrushedPath)
        Sw.Write(Xaml)
        Sw.Close()
    End Sub

    Private Sub lstBrushes_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles lstBrushes.SelectionChanged
        If lstBrushes.SelectedIndex = -1 Then Return
        Me.Brush = HatchBrushes.Clone(lstBrushes.SelectedItem)
        If TypeOf Me.Brush Is DrawingBrush Then
            TilePicker.SetParams()
        ElseIf TypeOf Me.Brush Is ImageBrush Then
            ImagePicker.SetParams()
        End If
    End Sub

    Private Sub BrushLabel_PreviewMouseDoubleClick(sender As Object, e As MouseButtonEventArgs)
        Me.DialogResult = True
        AddToRecent()
    End Sub

    Private Sub BtnReset_Click(sender As Object, e As RoutedEventArgs)
        lblBrushPreview.Background = Nothing
    End Sub
End Class
