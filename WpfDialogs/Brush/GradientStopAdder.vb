'''***************   NCore Softwares Pvt. Ltd., India   **************************
'''
'''   ColorPicker
'''
'''   Copyright (C) 2013 NCore Softwares Pvt. Ltd.
'''
'''   This program is provided to you under the terms of the Microsoft Public
'''   License (Ms-PL) as published at http://ColorPicker.codeplex.com/license
'''
'''**********************************************************************************

Imports System.Collections.ObjectModel

Friend Class GradientStopAdder
    Inherits Button

    Protected Overrides Sub OnPreviewMouseLeftButtonDown(ByVal e As MouseButtonEventArgs)
        MyBase.OnPreviewMouseLeftButtonDown(e)

        If TypeOf e.Source Is GradientStopAdder AndAlso Me.ColorPicker IsNot Nothing Then
            Dim btn As Button = TryCast(e.Source, Button)

            Dim _gs As New GradientStop()
            _gs.Offset = Mouse.GetPosition(btn).X / btn.ActualWidth
            _gs.Color = GetColorFromImage(e.GetPosition(Me))
            Me.ColorPicker.Gradients.Add(_gs)
            Me.ColorPicker.SelectedGradient = _gs
            Me.ColorPicker.Color = _gs.Color
            Me.ColorPicker.SetBrush()
        End If
    End Sub

    Private Function GetColorFromImage(ByVal p As Point) As Color
        Try
            Dim bounds As Rect = VisualTreeHelper.GetDescendantBounds(Me)
            Dim rtb As New RenderTargetBitmap(
                CInt(bounds.Width),
                CInt(bounds.Height),
                96, 96,
                PixelFormats.Default
            )
            rtb.Render(Me)

            Dim arr() As Byte
            Dim png As New PngBitmapEncoder()
            png.Frames.Add(BitmapFrame.Create(rtb))

            Using stream = New System.IO.MemoryStream()
                png.Save(stream)
                arr = stream.ToArray()
            End Using

            Dim bitmap As BitmapSource = BitmapFrame.Create(New System.IO.MemoryStream(arr))
            Dim pixels(3) As Byte
            Dim cb As New CroppedBitmap(bitmap, New Int32Rect(CInt(Fix(p.X)), CInt(Fix(p.Y)), 1, 1))
            cb.CopyPixels(pixels, 4, 0)
            Return Color.FromArgb(pixels(3), pixels(2), pixels(1), pixels(0))

        Catch e1 As Exception
            Return Me.ColorPicker.Color
        End Try
    End Function

    Public Property ColorPicker() As ColorPicker
        Get
            Return CType(GetValue(ColorPickerProperty), ColorPicker)
        End Get
        Set(ByVal value As ColorPicker)
            SetValue(ColorPickerProperty, value)
        End Set
    End Property
    Public Shared ReadOnly ColorPickerProperty As DependencyProperty = DependencyProperty.Register("ColorPicker", GetType(ColorPicker), GetType(GradientStopAdder))
End Class
