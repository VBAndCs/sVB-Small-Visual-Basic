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


Friend Class AlphaSelector
    Inherits BaseSelector

    Public Property Alpha() As Double
        Get
            Return CDbl(GetValue(AlphaProperty))
        End Get
        Set(ByVal value As Double)
            SetValue(AlphaProperty, value)
        End Set
    End Property

    Public Shared ReadOnly AlphaProperty As DependencyProperty = DependencyProperty.Register("Alpha", GetType(Double), GetType(AlphaSelector), New FrameworkPropertyMetadata(1.0, New PropertyChangedCallback(AddressOf AlphaChanged), New CoerceValueCallback(AddressOf AlphaCoerce)))
    Public Shared Sub AlphaChanged(ByVal o As Object, ByVal e As DependencyPropertyChangedEventArgs)
        Dim h As AlphaSelector = CType(o, AlphaSelector)
        h.SetAlphaOffset()
        h.SetColor()
    End Sub
    Public Shared Function AlphaCoerce(ByVal d As DependencyObject, ByVal Brightness As Object) As Object
        Dim v As Double = CDbl(Brightness)
        If v < 0 Then
            Return 0.0
        End If
        If v > 1 Then
            Return 1.0
        End If
        Return v
    End Function


    Public Property AlphaOffset() As Double
        Get
            Return CDbl(GetValue(AlphaOffsetProperty))
        End Get
        Private Set(ByVal value As Double)
            SetValue(AlphaOffsetProperty, value)
        End Set
    End Property
    Public Shared ReadOnly AlphaOffsetProperty As DependencyProperty = DependencyProperty.Register("AlphaOffset", GetType(Double), GetType(AlphaSelector), New UIPropertyMetadata(0.0))


    Protected Overrides Sub OnMouseMove(ByVal e As MouseEventArgs)
        If e.LeftButton = MouseButtonState.Pressed Then
            Dim p As Point = e.GetPosition(Me)

            If Orientation = Orientation.Vertical Then
                Alpha = 1 - (p.Y / Me.ActualHeight)
            Else
                Alpha = 1 - (p.X / Me.ActualWidth)
            End If
        End If
        MyBase.OnMouseMove(e)
    End Sub

    Protected Overrides Sub OnMouseDown(ByVal e As MouseButtonEventArgs)
        If e.LeftButton = MouseButtonState.Pressed Then
            Dim p As Point = e.GetPosition(Me)

            If Orientation = Orientation.Vertical Then
                Alpha = 1 - (p.Y / Me.ActualHeight)
            Else
                Alpha = 1 - (p.X / Me.ActualWidth)
            End If
        End If
        Mouse.Capture(Me)
        MyBase.OnMouseMove(e)
    End Sub

    Protected Overrides Sub OnMouseUp(ByVal e As MouseButtonEventArgs)
        Me.ReleaseMouseCapture()
        MyBase.OnMouseUp(e)
    End Sub

    Protected Overrides Sub OnRender(ByVal dc As DrawingContext)
        Dim lb As New LinearGradientBrush()

        lb.StartPoint = New Point(0, 0)

        If Orientation = Orientation.Vertical Then
            lb.EndPoint = New Point(0, 1)
        Else
            lb.EndPoint = New Point(1, 0)
        End If

        lb.GradientStops.Add(New GradientStop(Color.FromArgb(&HFF, &H0, &H0, &H0), 0.0))
        lb.GradientStops.Add(New GradientStop(Color.FromArgb(&H0, &H0, &H0, &H0), 1.0))
        dc.DrawRectangle(lb, Nothing, New Rect(0, 0, ActualWidth, ActualHeight))
        SetAlphaOffset()

    End Sub

    Protected Overrides Function ArrangeOverride(ByVal finalSize As Size) As Size
        SetAlphaOffset()
        Return MyBase.ArrangeOverride(finalSize)
    End Function


    Private Sub SetAlphaOffset()
        Dim length As Double = ActualHeight
        If Orientation = Orientation.Horizontal Then
            length = ActualWidth
        End If
        AlphaOffset = length - (length * Alpha)
    End Sub

    Private Sub SetColor()
        Color = Color.FromArgb(CByte(Math.Round(Alpha * 255)), 0, 0, 0)
    End Sub
End Class
