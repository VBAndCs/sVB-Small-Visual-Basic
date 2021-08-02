Imports System.Windows.Controls.Primitives
Imports System.ComponentModel

Public Class ZoomBox
    Inherits Control

    Dim ScrollViewer As ScrollViewer
    Private zoomThumb As Thumb
    Private zoomCanvas As Canvas
    Friend zoomSlider As Slider
    Private ZoomExpander As Expander
    Private Scale As Double = 1.0
    Private WithEvents DesignerCanvas As Canvas
    Dim ExitScrollChanged As Boolean

    Public Sub New()
        Dim resourceLocater As Uri = New Uri("/DiagramHelper;component/Resources/zoomboxdecorator.xaml", System.UriKind.Relative)
        Dim ResDec As ResourceDictionary = Application.LoadComponent(resourceLocater)
        Me.Resources.MergedDictionaries.Add(ResDec)
        Me.Style = FindResource("ZoomBoxStyle")
    End Sub

    Private Sub ZoomBox_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        Me.DesignerCanvas = Designer.DesignerCanvas
        Me.zoomThumb = TryCast(Template.FindName("PART_ZoomThumb", Me), Thumb)
        Me.zoomCanvas = TryCast(Template.FindName("PART_ZoomCanvas", Me), Canvas)
        Me.zoomSlider = TryCast(Template.FindName("PART_ZoomSlider", Me), Slider)
        Me.ZoomExpander = TryCast(Template.FindName("PART_ZoomExpander", Me), Expander)

        If Designer IsNot Nothing Then Designer.ZoomBox = Me

        AddHandler zoomThumb.DragDelta, AddressOf Thumb_DragDelta
        AddHandler zoomSlider.ValueChanged, AddressOf ZoomSlider_ValueChanged
    End Sub

    Public Property Designer As Designer
        Get
            Return GetValue(DesignerProperty)
        End Get

        Set(ByVal value As Designer)
            SetValue(DesignerProperty, value)
            Designer.ZoomBox = Me

            Refresh()

        End Set
    End Property

    Public Shared ReadOnly DesignerProperty As DependencyProperty = _
                           DependencyProperty.Register("Designer", _
                           GetType(Designer), GetType(ZoomBox))


    Private Sub ZoomSlider_ValueChanged(ByVal sender As Object, ByVal e As RoutedPropertyChangedEventArgs(Of Double))
        If ScrollViewer Is Nothing Then Return
        Dim Xof = ScrollViewer.HorizontalOffset
        Dim Yof = ScrollViewer.VerticalOffset

        Me.ExitScrollChanged = True

        Me.Scale = e.NewValue / 100
        Me.Designer.Scale = Me.Scale

        Dim r = e.NewValue / e.OldValue

        Xof = Xof * r + ScrollViewer.ViewportWidth * (r - 1) / 2
        ScrollViewer.ScrollToHorizontalOffset(Xof)

        Yof = Yof * r + ScrollViewer.ViewportHeight * (r - 1) / 2
        ScrollViewer.ScrollToVerticalOffset(Yof)
        Me.ExitScrollChanged = False

    End Sub

    Private Sub Thumb_DragDelta(ByVal sender As Object, ByVal e As DragDeltaEventArgs)
        Dim scale, xOffset, yOffset As Double
        Me.InvalidateScale(scale, xOffset, yOffset)

        Me.ScrollViewer.ScrollToHorizontalOffset(Me.ScrollViewer.HorizontalOffset + e.HorizontalChange / scale)
        Me.ScrollViewer.ScrollToVerticalOffset(Me.ScrollViewer.VerticalOffset + e.VerticalChange / scale)
    End Sub


    Private Sub InvalidateScale(ByRef scale As Double, ByRef xOffset As Double, ByRef yOffset As Double)
        Dim scaleX, scaleY As Double

        If Me.DesignerCanvas Is Nothing Then Return

        ' designer canvas size
        Dim w As Double = Me.DesignerCanvas.ActualWidth * Me.Scale
        Dim h As Double = Me.DesignerCanvas.ActualHeight * Me.Scale

        ' zoom canvas size
        Dim x = ZoomExpander.ActualWidth - 10
        Dim y As Double = Me.zoomCanvas.ActualHeight


        scaleX = x / w
        scaleY = y / h

        scale = If(scaleX < scaleY, scaleX, scaleY)

        xOffset = (x - scale * w) / 2
        If xOffset < 0 Then xOffset = 0
        yOffset = (y - scale * h) / 2
        If yOffset < 0 Then yOffset = 0
    End Sub

    Private Sub ScrollViewer_ScrollChanged(sender As Object, e As ScrollChangedEventArgs)
        If ExitScrollChanged Then Return
        If ScrollViewer Is Nothing Then
            Refresh()
            Exit Sub
        End If

        ExitScrollChanged = True

        Dim scale, xOffset, yOffset As Double
        Me.InvalidateScale(scale, xOffset, yOffset)
        If scale = 0 Then Return

        Dim Left = xOffset + Me.ScrollViewer.HorizontalOffset * scale
        Canvas.SetLeft(Me.zoomThumb, Left)
        Dim Top = yOffset + Me.ScrollViewer.VerticalOffset * scale
        Canvas.SetTop(Me.zoomThumb, Top)

        Dim ThumbWidth = Me.ScrollViewer.ViewportWidth * scale
        Dim ViewWidth = ZoomExpander.ActualWidth - 10 - 2 * xOffset
        If ThumbWidth > ViewWidth Then
            Me.zoomThumb.Width = ViewWidth
        Else
            Me.zoomThumb.Width = ThumbWidth
        End If

        Dim ThumbHeight = Me.ScrollViewer.ViewportHeight * scale
        Dim ViewHeight = zoomCanvas.ActualHeight - 2 * yOffset
        If ThumbHeight > ViewHeight Then
            Me.zoomThumb.Height = ViewHeight
        Else
            Me.zoomThumb.Height = ThumbHeight
        End If

        ExitScrollChanged = False
    End Sub

    Public Sub Refresh()
        ScrollViewer = Designer.ScrollViewer
        If ScrollViewer IsNot Nothing Then
            RemoveHandler ScrollViewer.ScrollChanged, AddressOf ScrollViewer_ScrollChanged
            AddHandler ScrollViewer.ScrollChanged, AddressOf ScrollViewer_ScrollChanged
            ScrollViewer_ScrollChanged(Nothing, Nothing)
        End If
    End Sub
End Class
