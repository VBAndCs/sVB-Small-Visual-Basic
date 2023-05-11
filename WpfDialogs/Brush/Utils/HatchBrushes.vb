Public Class HatchBrushes

    Public Shared Function GetAllHatchBrushes() As List(Of DrawingBrush)
        Dim Lst As New List(Of DrawingBrush)
        For Each HStyle As HatchStyle In [Enum].GetValues(GetType(HatchStyle))
            Dim b As New DrawingBrush
            SetHatchBrushParam(b, New HatchParams(HStyle, Brushes.White, Brushes.Black, 1, Nothing, Nothing))
            b.Transform = New ScaleTransform(2, 2)
            Lst.Add(b)
        Next
        Return Lst
    End Function

    Public Shared Function GetHatchStyle(ByVal element As DependencyObject) As HatchStyle
        If element Is Nothing Then
            Throw New ArgumentNullException("element")
        End If

        Return element.GetValue(HatchStyleProperty)
    End Function

    Public Shared Sub SetHatchStyle(ByVal element As DependencyObject, ByVal value As HatchStyle)
        If element Is Nothing Then
            Throw New ArgumentNullException("element")
        End If

        element.SetValue(HatchStyleProperty, value)
    End Sub

    Public Shared ReadOnly HatchStyleProperty As  _
                           DependencyProperty = DependencyProperty.RegisterAttached("HatchStyle", _
                           GetType(HatchStyle), GetType(HatchBrushes), _
                           New PropertyMetadata(AddressOf OnHatchChanged))

    Public Shared Function GetHatchBrushParam(ByVal element As DependencyObject) As HatchParams
        If element Is Nothing Then
            Throw New ArgumentNullException("element")
        End If

        Return New HatchParams(GetHatchStyle(element),
                                                  GetBackground(element),
                                                  GetForeground(element),
                                                  GetPenThickness(element),
                                                  GetHatchTransform(element),
                                                  GetShapeTransform(element))
    End Function

    Public Shared Sub SetHatchBrushParam(ByVal Brush As DrawingBrush, ByVal Params As HatchParams)
        If DontChange Then Return
        If Brush Is Nothing Then Return

        DontChange = True
        SetHatchStyle(Brush, Params.HatchStyle)
        SetBackground(Brush, Params.Background)
        SetForeground(Brush, Params.Foreground)
        SetPenThickness(Brush, Params.PenThickness)
        SetHatchTransform(Brush, Params.HatchTransform)
        SetShapeTransform(Brush, Params.ShapeTransform)
        DontChange = False
        OnHatchChanged(Brush, Nothing)
    End Sub

    Public Shared Function GetHatchTransform(ByVal element As DependencyObject) As Transform
        If element Is Nothing Then
            Throw New ArgumentNullException("element")
        End If

        Return element.GetValue(HatchTransformProperty)
    End Function

    Public Shared Sub SetHatchTransform(ByVal element As DependencyObject, ByVal value As Transform)
        If element Is Nothing Then
            Throw New ArgumentNullException("element")
        End If

        element.SetValue(HatchTransformProperty, value)
    End Sub

    Public Shared ReadOnly HatchTransformProperty As  _
                           DependencyProperty = DependencyProperty.RegisterAttached("HatchTransform", _
                           GetType(Transform), GetType(HatchBrushes), _
                           New PropertyMetadata(AddressOf OnHatchChanged))

    Public Shared Function GetShapeTransform(ByVal element As DependencyObject) As TransformGroup
        If element Is Nothing Then
            Throw New ArgumentNullException("element")
        End If

        Return element.GetValue(ShapeTransformProperty)
    End Function

    Public Shared Sub SetShapeTransform(ByVal element As DependencyObject, ByVal value As Transform)
        If element Is Nothing Then
            Throw New ArgumentNullException("element")
        End If

        element.SetValue(ShapeTransformProperty, value)
    End Sub

    Public Shared ReadOnly ShapeTransformProperty As  _
                           DependencyProperty = DependencyProperty.RegisterAttached("ShapeTransform", _
                           GetType(Transform), GetType(HatchBrushes), _
                           New PropertyMetadata(AddressOf OnHatchChanged))

    Public Shared Function GetBackground(ByVal element As DependencyObject) As Brush
        If element Is Nothing Then
            Throw New ArgumentNullException("element")
        End If

        Return element.GetValue(BackgroundProperty)
    End Function

    Public Shared Sub SetBackground(ByVal element As DependencyObject, ByVal value As Brush)
        If element Is Nothing Then
            Throw New ArgumentNullException("element")
        End If

        element.SetValue(BackgroundProperty, value)
    End Sub

    Public Shared ReadOnly BackgroundProperty As  _
                           DependencyProperty = DependencyProperty.RegisterAttached("Background", _
                           GetType(Brush), GetType(HatchBrushes), _
                           New PropertyMetadata(Brushes.White, AddressOf OnHatchChanged))

    Public Shared Function GetForeground(ByVal element As DependencyObject) As Brush
        If element Is Nothing Then
            Throw New ArgumentNullException("element")
        End If

        Return element.GetValue(ForegroundProperty)
    End Function

    Public Shared Sub SetForeground(ByVal element As DependencyObject, ByVal value As Brush)
        If element Is Nothing Then
            Throw New ArgumentNullException("element")
        End If

        element.SetValue(ForegroundProperty, value)
    End Sub

    Public Shared ReadOnly ForegroundProperty As  _
                           DependencyProperty = DependencyProperty.RegisterAttached("Foreground", _
                           GetType(Brush), GetType(HatchBrushes), _
                           New PropertyMetadata(Brushes.Black, AddressOf OnHatchChanged))

    Public Shared Function GetPenThickness(ByVal element As DependencyObject) As Double
        If element Is Nothing Then
            Throw New ArgumentNullException("element")
        End If

        Return element.GetValue(PenThicknessProperty)
    End Function

    Public Shared Sub SetPenThickness(element As DependencyObject, ByVal value As Double)
        If element Is Nothing Then
            Throw New ArgumentNullException("element")
        End If

        element.SetValue(PenThicknessProperty, value)
    End Sub

    Public Shared ReadOnly PenThicknessProperty As  _
                           DependencyProperty = DependencyProperty.RegisterAttached("PenThickness", _
                           GetType(Double), GetType(HatchBrushes), _
                           New PropertyMetadata(1.0, AddressOf OnHatchChanged))

    Shared DontChange As Boolean = False

    Shared Sub OnHatchChanged(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
        If DontChange Then Return
        Dim DBrush = TryCast(d, DrawingBrush)
        If DBrush Is Nothing Then Return

        ModifyHatchBrush(DBrush)
    End Sub

    Public Shared Function GetHatchWidth(ByVal element As DependencyObject) As Double
        If element Is Nothing Then
            Throw New ArgumentNullException("element")
        End If

        Return element.GetValue(HatchWidthProperty)
    End Function

    Public Shared Sub SetHatchWidth(ByVal element As DependencyObject, ByVal value As Double)
        If element Is Nothing Then
            Throw New ArgumentNullException("element")
        End If

        element.SetValue(HatchWidthProperty, value)
    End Sub

    Public Shared ReadOnly HatchWidthProperty As  _
                           DependencyProperty = DependencyProperty.RegisterAttached("HatchWidth", _
                           GetType(Double), GetType(HatchBrushes))

    Public Shared Function GetHatchHeight(ByVal element As DependencyObject) As Double
        If element Is Nothing Then
            Throw New ArgumentNullException("element")
        End If

        Return element.GetValue(HatchHeightProperty)
    End Function

    Public Shared Sub SetHatchHeight(ByVal element As DependencyObject, ByVal value As Double)
        If element Is Nothing Then
            Throw New ArgumentNullException("element")
        End If

        element.SetValue(HatchHeightProperty, value)
    End Sub

    Public Shared ReadOnly HatchHeightProperty As  _
                           DependencyProperty = DependencyProperty.RegisterAttached("HatchHeight", _
                           GetType(Double), GetType(HatchBrushes))

    Shared Function Clone(brush As Brush) As Brush
        If DontChange OrElse brush Is Nothing Then Return Nothing

        DontChange = True
        Dim ClonedBrush = brush.CloneCurrentValue
        Dim HatchBrush = TryCast(brush, DrawingBrush)
        If HatchBrush IsNot Nothing Then
            Dim B = CType(ClonedBrush, DrawingBrush)
            B.Viewbox = HatchBrush.Viewbox
            B.Viewport = HatchBrush.Viewport
            B.TileMode = HatchBrush.TileMode
            B.Stretch = HatchBrush.Stretch
        End If
        DontChange = False
        Return ClonedBrush
    End Function

    Shared Function CloneDrawingBrush(brush As Brush) As Brush
        If Not TypeOf brush Is DrawingBrush Then Return brush

        Dim Dbrush As DrawingBrush = brush
        Dim B As New DrawingBrush
        If B.ViewboxUnits <> Dbrush.ViewboxUnits Then B.ViewboxUnits = Dbrush.ViewboxUnits
        If B.Viewbox <> Dbrush.Viewbox Then B.Viewbox = Dbrush.Viewbox
        If B.ViewportUnits <> Dbrush.ViewportUnits Then B.ViewportUnits = Dbrush.ViewportUnits
        If B.Viewport <> Dbrush.Viewport Then B.Viewport = Dbrush.Viewport
        If Dbrush.RelativeTransform.Value <> Matrix.Identity Then B.RelativeTransform = Dbrush.RelativeTransform
        If B.TileMode <> Dbrush.TileMode Then B.TileMode = Dbrush.TileMode
        If B.Stretch <> Dbrush.Stretch Then B.Stretch = Dbrush.Stretch
        If B.AlignmentX <> Dbrush.AlignmentX Then B.AlignmentX = Dbrush.AlignmentX
        If B.AlignmentY <> Dbrush.AlignmentY Then B.AlignmentY = Dbrush.AlignmentY
        B.Drawing = Dbrush.Drawing.CloneCurrentValue
        Return B
    End Function

#Region "Draw Hatch"

    Private mp_oGraphics As DrawingContext

    Private Enum HatchType
        Rectangle = 1
        Line = 2
    End Enum

    Private Shared Sub ModifyHatchBrush(DBrush As DrawingBrush)
        Dim HatchStyle As HatchStyle = GetHatchStyle(DBrush)
        Dim Background As Brush = GetBackground(DBrush)
        Dim Foreground As Brush = GetForeground(DBrush)
        Dim PenThickness As Double = GetPenThickness(DBrush)

        Dim oHatchGroup As New GeometryGroup()        
        Dim oHatchCtrlGroup As New GeometryGroup()
        Dim lWidth As Double = 0
        Dim lHeight As Double = 0
        Dim yType As HatchType = HatchType.Line
        Dim bAliased As Boolean = False
        Select Case HatchStyle
            Case HatchStyle.Horizontal
                lWidth = 8
                lHeight = 8
                oHatchGroup.Children.Add(mp_GLine(0, 0, 7, 0))
                yType = HatchType.Line
            Case HatchStyle.Vertical
                lWidth = 8
                lHeight = 8
                oHatchGroup.Children.Add(mp_GLine(0, 0, 0, 7))
                yType = HatchType.Line
            Case HatchStyle.ForwardDiagonal
                lWidth = 16
                lHeight = 16
                oHatchGroup.Children.Add(mp_GLine(0, 12, 3, 15))
                oHatchGroup.Children.Add(mp_GLine(0, 4, 11, 15))
                oHatchGroup.Children.Add(mp_GLine(4, 0, 15, 11))
                oHatchGroup.Children.Add(mp_GLine(12, 0, 15, 3))
                System.Diagnostics.Debug.Write(oHatchGroup)
                yType = HatchType.Line
            Case HatchStyle.BackwardDiagonal
                lWidth = 16
                lHeight = 16
                oHatchGroup.Children.Add(mp_GLine(4, 16, 15, 3))
                oHatchGroup.Children.Add(mp_GLine(12, 16, 15, 11))
                oHatchGroup.Children.Add(mp_GLine(0, 4, 3, -1))
                oHatchGroup.Children.Add(mp_GLine(12, 0, -1, 11))
                yType = HatchType.Line
            Case HatchStyle.LargeGrid
                lWidth = 8
                lHeight = 8
                oHatchGroup.Children.Add(mp_GLine(0, 0, 7, 0))
                oHatchGroup.Children.Add(mp_GLine(0, 0, 0, 7))
                yType = HatchType.Line
            Case HatchStyle.DiagonalCross
                lWidth = 8
                lHeight = 8
                oHatchGroup.Children.Add(mp_GLine(4, 0, -1, 3))
                oHatchGroup.Children.Add(mp_GLine(0, 4, 3, 7))
                oHatchGroup.Children.Add(mp_GLine(4, 8, 7, 3))
                oHatchGroup.Children.Add(mp_GLine(8, 4, 3, -1))
                yType = HatchType.Line
                bAliased = False
            Case HatchStyle.Percent05
                lWidth = 8
                lHeight = 8
                oHatchGroup.Children.Add(mp_GPoint(0, 4))
                oHatchGroup.Children.Add(mp_GPoint(4, 0))
                yType = HatchType.Line
            Case HatchStyle.Percent10
                lWidth = 8
                lHeight = 8
                oHatchGroup.Children.Add(mp_GPoint(0, 2))
                oHatchGroup.Children.Add(mp_GPoint(4, 0))
                oHatchGroup.Children.Add(mp_GPoint(0, 6))
                oHatchGroup.Children.Add(mp_GPoint(4, 4))
                yType = HatchType.Line
            Case HatchStyle.Percent20
                lWidth = 4
                lHeight = 4
                oHatchGroup.Children.Add(mp_GPoint(0, 0))
                oHatchGroup.Children.Add(mp_GPoint(2, 2))
                yType = HatchType.Line
            Case HatchStyle.Percent25
                lWidth = 4
                lHeight = 4
                oHatchGroup.Children.Add(mp_GPoint(0, 0))
                oHatchGroup.Children.Add(mp_GPoint(0, 2))
                oHatchGroup.Children.Add(mp_GPoint(2, 1))
                oHatchGroup.Children.Add(mp_GPoint(2, 3))
                yType = HatchType.Line
            Case HatchStyle.Percent30
                lWidth = 4
                lHeight = 4
                oHatchGroup.Children.Add(mp_GPoint(0, 0))
                oHatchGroup.Children.Add(mp_GPoint(1, 1))
                oHatchGroup.Children.Add(mp_GPoint(2, 2))
                oHatchGroup.Children.Add(mp_GPoint(3, 3))
                oHatchGroup.Children.Add(mp_GPoint(2, 0))
                oHatchGroup.Children.Add(mp_GPoint(0, 2))
                yType = HatchType.Line
            Case HatchStyle.Percent40
                lWidth = 8
                lHeight = 8
                oHatchGroup.Children.Add(mp_GPoint(0, 0))
                oHatchGroup.Children.Add(mp_GPoint(2, 0))
                oHatchGroup.Children.Add(mp_GPoint(4, 0))
                oHatchGroup.Children.Add(mp_GPoint(6, 0))

                oHatchGroup.Children.Add(mp_GPoint(3, 1))
                oHatchGroup.Children.Add(mp_GPoint(5, 1))
                oHatchGroup.Children.Add(mp_GPoint(7, 1))

                oHatchGroup.Children.Add(mp_GPoint(0, 2))
                oHatchGroup.Children.Add(mp_GPoint(2, 2))
                oHatchGroup.Children.Add(mp_GPoint(4, 2))
                oHatchGroup.Children.Add(mp_GPoint(6, 2))

                oHatchGroup.Children.Add(mp_GPoint(1, 3))
                oHatchGroup.Children.Add(mp_GPoint(3, 3))
                oHatchGroup.Children.Add(mp_GPoint(5, 3))
                oHatchGroup.Children.Add(mp_GPoint(7, 3))

                oHatchGroup.Children.Add(mp_GPoint(0, 4))
                oHatchGroup.Children.Add(mp_GPoint(2, 4))
                oHatchGroup.Children.Add(mp_GPoint(4, 4))
                oHatchGroup.Children.Add(mp_GPoint(6, 4))
                oHatchGroup.Children.Add(mp_GPoint(1, 5))
                oHatchGroup.Children.Add(mp_GPoint(3, 5))
                oHatchGroup.Children.Add(mp_GPoint(7, 5))
                oHatchGroup.Children.Add(mp_GPoint(0, 6))
                oHatchGroup.Children.Add(mp_GPoint(2, 6))
                oHatchGroup.Children.Add(mp_GPoint(4, 6))
                oHatchGroup.Children.Add(mp_GPoint(6, 6))
                oHatchGroup.Children.Add(mp_GPoint(1, 7))
                oHatchGroup.Children.Add(mp_GPoint(3, 7))
                oHatchGroup.Children.Add(mp_GPoint(5, 7))
                oHatchGroup.Children.Add(mp_GPoint(7, 7))
                yType = HatchType.Line
            Case HatchStyle.Percent50
                lWidth = 2
                lHeight = 2
                oHatchGroup.Children.Add(mp_GPoint(0, 0))
                oHatchGroup.Children.Add(mp_GPoint(1, 1))
                yType = HatchType.Line
            Case HatchStyle.Percent60
                lWidth = 4
                lHeight = 4
                oHatchGroup.Children.Add(mp_GPoint(0, 0))
                oHatchGroup.Children.Add(mp_GPoint(2, 0))
                oHatchGroup.Children.Add(mp_GPoint(3, 0))
                oHatchGroup.Children.Add(mp_GPoint(1, 1))
                oHatchGroup.Children.Add(mp_GPoint(3, 1))
                oHatchGroup.Children.Add(mp_GPoint(0, 2))
                oHatchGroup.Children.Add(mp_GPoint(1, 2))
                oHatchGroup.Children.Add(mp_GPoint(2, 2))
                oHatchGroup.Children.Add(mp_GPoint(1, 3))
                oHatchGroup.Children.Add(mp_GPoint(3, 3))
                yType = HatchType.Line
            Case HatchStyle.Percent70
                lWidth = 4
                lHeight = 4
                oHatchGroup.Children.Add(mp_GPoint(1, 0))
                oHatchGroup.Children.Add(mp_GPoint(2, 0))
                oHatchGroup.Children.Add(mp_GPoint(3, 0))
                oHatchGroup.Children.Add(mp_GPoint(0, 1))
                oHatchGroup.Children.Add(mp_GPoint(1, 1))
                oHatchGroup.Children.Add(mp_GPoint(3, 1))
                oHatchGroup.Children.Add(mp_GPoint(1, 2))
                oHatchGroup.Children.Add(mp_GPoint(2, 2))
                oHatchGroup.Children.Add(mp_GPoint(3, 2))
                oHatchGroup.Children.Add(mp_GPoint(0, 3))
                oHatchGroup.Children.Add(mp_GPoint(1, 3))
                oHatchGroup.Children.Add(mp_GPoint(3, 3))
                yType = HatchType.Line
            Case HatchStyle.Percent75
                lWidth = 4
                lHeight = 4
                oHatchGroup.Children.Add(mp_GPoint(2, 0))
                oHatchGroup.Children.Add(mp_GPoint(0, 2))
                InvertColors(Foreground, Background)
                yType = HatchType.Line
            Case HatchStyle.Percent80
                lWidth = 8
                lHeight = 7
                oHatchGroup.Children.Add(mp_GPoint(3, 0))
                oHatchGroup.Children.Add(mp_GPoint(3, 4))
                oHatchGroup.Children.Add(mp_GPoint(7, 2))
                oHatchGroup.Children.Add(mp_GPoint(7, 6))
                InvertColors(Foreground, Background)
                yType = HatchType.Line
            Case HatchStyle.Percent90
                lWidth = 8
                lHeight = 8
                oHatchGroup.Children.Add(mp_GPoint(0, 5))
                oHatchGroup.Children.Add(mp_GPoint(4, 1))
                InvertColors(Foreground, Background)
                yType = HatchType.Line
            Case HatchStyle.LightDownwardDiagonal
                lWidth = 4
                lHeight = 4
                oHatchGroup.Children.Add(mp_GPoint(0, 0))
                oHatchGroup.Children.Add(mp_GPoint(1, 1))
                oHatchGroup.Children.Add(mp_GPoint(2, 2))
                oHatchGroup.Children.Add(mp_GPoint(3, 3))
                yType = HatchType.Line
            Case HatchStyle.LightUpwardDiagonal
                lWidth = 4
                lHeight = 4
                oHatchGroup.Children.Add(mp_GPoint(0, 3))
                oHatchGroup.Children.Add(mp_GPoint(1, 2))
                oHatchGroup.Children.Add(mp_GPoint(2, 1))
                oHatchGroup.Children.Add(mp_GPoint(3, 0))
                yType = HatchType.Line
            Case HatchStyle.DarkDownwardDiagonal
                lWidth = 4
                lHeight = 4
                oHatchGroup.Children.Add(mp_GPoint(0, 3))
                oHatchGroup.Children.Add(mp_GPoint(0, 0))
                oHatchGroup.Children.Add(mp_GPoint(1, 1))
                oHatchGroup.Children.Add(mp_GPoint(2, 2))
                oHatchGroup.Children.Add(mp_GPoint(3, 3))
                oHatchGroup.Children.Add(mp_GPoint(1, 0))
                oHatchGroup.Children.Add(mp_GPoint(2, 1))
                oHatchGroup.Children.Add(mp_GPoint(3, 2))
                yType = HatchType.Line
            Case HatchStyle.DarkUpwardDiagonal
                lWidth = 4
                lHeight = 4
                oHatchGroup.Children.Add(mp_GPoint(0, 0))
                oHatchGroup.Children.Add(mp_GPoint(0, 1))
                oHatchGroup.Children.Add(mp_GPoint(1, 0))
                oHatchGroup.Children.Add(mp_GPoint(1, 3))
                oHatchGroup.Children.Add(mp_GPoint(2, 2))
                oHatchGroup.Children.Add(mp_GPoint(3, 1))
                oHatchGroup.Children.Add(mp_GPoint(2, 3))
                oHatchGroup.Children.Add(mp_GPoint(3, 2))
                yType = HatchType.Line
            Case HatchStyle.WideDownwardDiagonal
                lWidth = 8
                lHeight = 8
                oHatchGroup.Children.Add(mp_GLine(3, 0, 5, 0))
                oHatchGroup.Children.Add(mp_GLine(4, 1, 6, 1))
                oHatchGroup.Children.Add(mp_GLine(5, 2, 7, 2))
                oHatchGroup.Children.Add(mp_GPoint(0, 3))
                oHatchGroup.Children.Add(mp_GLine(6, 3, 7, 3))
                oHatchGroup.Children.Add(mp_GLine(0, 4, 1, 4))
                oHatchGroup.Children.Add(mp_GPoint(7, 4))
                oHatchGroup.Children.Add(mp_GLine(0, 5, 2, 5))
                oHatchGroup.Children.Add(mp_GLine(1, 6, 3, 6))
                oHatchGroup.Children.Add(mp_GLine(2, 7, 4, 7))
                oHatchCtrlGroup.Children.Add(mp_GPoint(7, 7))
                yType = HatchType.Line
            Case HatchStyle.WideUpwardDiagonal
                lWidth = 8
                lHeight = 8
                oHatchGroup.Children.Add(mp_GLine(4, 0, 6, 0))
                oHatchGroup.Children.Add(mp_GLine(3, 1, 5, 1))
                oHatchGroup.Children.Add(mp_GLine(2, 2, 4, 2))
                oHatchGroup.Children.Add(mp_GLine(1, 3, 3, 3))
                oHatchGroup.Children.Add(mp_GLine(0, 4, 2, 4))
                oHatchGroup.Children.Add(mp_GLine(0, 5, 1, 5))
                oHatchGroup.Children.Add(mp_GPoint(7, 5))
                oHatchGroup.Children.Add(mp_GPoint(0, 6))
                oHatchGroup.Children.Add(mp_GLine(6, 6, 7, 6))
                oHatchGroup.Children.Add(mp_GLine(5, 7, 7, 7))
                oHatchCtrlGroup.Children.Add(mp_GLine(0, 0, 7, 7))
                yType = HatchType.Line
            Case HatchStyle.LightVertical
                lWidth = 4
                lHeight = 4
                oHatchGroup.Children.Add(mp_GLine(0, 0, 0, 3))
                yType = HatchType.Line
            Case HatchStyle.LightHorizontal
                lWidth = 4
                lHeight = 4
                oHatchGroup.Children.Add(mp_GLine(0, 0, 3, 0))
                yType = HatchType.Line
            Case HatchStyle.NarrowVertical
                lWidth = 2
                lHeight = 2
                oHatchGroup.Children.Add(mp_GLine(1, 0, 1, 1))
                yType = HatchType.Line
            Case HatchStyle.NarrowHorizontal
                lWidth = 2
                lHeight = 2
                oHatchGroup.Children.Add(mp_GLine(0, 1, 1, 1))
                yType = HatchType.Line
            Case HatchStyle.DarkVertical
                lWidth = 4
                lHeight = 4
                oHatchGroup.Children.Add(mp_GLine(0, 0, 0, 3))
                oHatchGroup.Children.Add(mp_GLine(1, 0, 1, 3))
                yType = HatchType.Line
            Case HatchStyle.DarkHorizontal
                lWidth = 4
                lHeight = 4
                oHatchGroup.Children.Add(mp_GLine(0, 0, 3, 0))
                oHatchGroup.Children.Add(mp_GLine(0, 1, 3, 1))
                yType = HatchType.Line
            Case HatchStyle.DashedDownwardDiagonal
                lWidth = 8
                lHeight = 8
                oHatchGroup.Children.Add(mp_GPoint(0, 2))
                oHatchGroup.Children.Add(mp_GPoint(1, 3))
                oHatchGroup.Children.Add(mp_GPoint(2, 4))
                oHatchGroup.Children.Add(mp_GPoint(3, 5))
                oHatchGroup.Children.Add(mp_GPoint(4, 2))
                oHatchGroup.Children.Add(mp_GPoint(5, 3))
                oHatchGroup.Children.Add(mp_GPoint(6, 4))
                oHatchGroup.Children.Add(mp_GPoint(7, 5))
                yType = HatchType.Line
            Case HatchStyle.DashedUpwardDiagonal
                lWidth = 8
                lHeight = 8
                oHatchGroup.Children.Add(mp_GPoint(0, 7))
                oHatchGroup.Children.Add(mp_GPoint(1, 6))
                oHatchGroup.Children.Add(mp_GPoint(2, 5))
                oHatchGroup.Children.Add(mp_GPoint(3, 4))
                oHatchGroup.Children.Add(mp_GPoint(4, 7))
                oHatchGroup.Children.Add(mp_GPoint(5, 6))
                oHatchGroup.Children.Add(mp_GPoint(6, 5))
                oHatchGroup.Children.Add(mp_GPoint(7, 4))
                yType = HatchType.Line
            Case HatchStyle.DashedHorizontal
                lWidth = 8
                lHeight = 8
                oHatchGroup.Children.Add(mp_GLine(4, 0, 7, 0))
                oHatchGroup.Children.Add(mp_GLine(0, 4, 3, 4))
                oHatchCtrlGroup.Children.Add(mp_GPoint(7, 7))
                yType = HatchType.Line
            Case HatchStyle.DashedVertical
                lWidth = 8
                lHeight = 8
                oHatchGroup.Children.Add(mp_GLine(0, 0, 0, 1))
                oHatchGroup.Children.Add(mp_GLine(0, 6, 0, 7))
                oHatchGroup.Children.Add(mp_GLine(4, 2, 4, 5))
                oHatchCtrlGroup.Children.Add(mp_GPoint(7, 7))
                yType = HatchType.Line
            Case HatchStyle.SmallConfetti
                lWidth = 8
                lHeight = 8
                oHatchGroup.Children.Add(mp_GPoint(0, 0))
                oHatchGroup.Children.Add(mp_GPoint(4, 1))
                oHatchGroup.Children.Add(mp_GPoint(1, 2))
                oHatchGroup.Children.Add(mp_GPoint(6, 3))
                oHatchGroup.Children.Add(mp_GPoint(3, 4))
                oHatchGroup.Children.Add(mp_GPoint(7, 5))
                oHatchGroup.Children.Add(mp_GPoint(2, 6))
                oHatchGroup.Children.Add(mp_GPoint(5, 7))
                oHatchCtrlGroup.Children.Add(mp_GPoint(7, 7))
                yType = HatchType.Line
            Case HatchStyle.LargeConfetti
                lWidth = 8
                lHeight = 8
                oHatchGroup.Children.Add(mp_GLine(0, 1, 0, 2))
                oHatchGroup.Children.Add(mp_GLine(0, 6, 0, 7))
                oHatchGroup.Children.Add(mp_GLine(1, 6, 1, 7))
                oHatchGroup.Children.Add(mp_GLine(2, 2, 2, 3))
                oHatchGroup.Children.Add(mp_GLine(3, 2, 3, 3))
                oHatchGroup.Children.Add(mp_GLine(3, 5, 3, 6))
                oHatchGroup.Children.Add(mp_GLine(4, 0, 4, 1))
                oHatchGroup.Children.Add(mp_GLine(4, 5, 4, 6))
                oHatchGroup.Children.Add(mp_GLine(5, 0, 5, 1))
                oHatchGroup.Children.Add(mp_GLine(6, 4, 6, 5))
                oHatchGroup.Children.Add(mp_GLine(7, 1, 7, 2))
                oHatchGroup.Children.Add(mp_GLine(7, 4, 7, 5))
                oHatchCtrlGroup.Children.Add(mp_GPoint(7, 7))
                yType = HatchType.Line
            Case HatchStyle.Zigzag
                lWidth = 8
                lHeight = 8
                oHatchGroup.Children.Add(mp_GPoint(0, 0))
                oHatchGroup.Children.Add(mp_GPoint(1, 1))
                oHatchGroup.Children.Add(mp_GPoint(2, 2))
                oHatchGroup.Children.Add(mp_GLine(3, 3, 4, 3))
                oHatchGroup.Children.Add(mp_GPoint(5, 2))
                oHatchGroup.Children.Add(mp_GPoint(6, 1))
                oHatchGroup.Children.Add(mp_GPoint(7, 0))

                oHatchGroup.Children.Add(mp_GPoint(0, 4))
                oHatchGroup.Children.Add(mp_GPoint(1, 5))
                oHatchGroup.Children.Add(mp_GPoint(2, 6))
                oHatchGroup.Children.Add(mp_GLine(3, 7, 4, 7))
                oHatchGroup.Children.Add(mp_GPoint(5, 6))
                oHatchGroup.Children.Add(mp_GPoint(6, 5))
                oHatchGroup.Children.Add(mp_GPoint(7, 4))
                yType = HatchType.Line
            Case HatchStyle.Wave
                lWidth = 8
                lHeight = 8
                oHatchGroup.Children.Add(mp_GPoint(2, 0))
                oHatchGroup.Children.Add(mp_GPoint(5, 0))
                oHatchGroup.Children.Add(mp_GPoint(7, 0))
                oHatchGroup.Children.Add(mp_GLine(0, 2, 1, 2))

                oHatchGroup.Children.Add(mp_GLine(3, 4, 4, 4))
                oHatchGroup.Children.Add(mp_GPoint(2, 4))
                oHatchGroup.Children.Add(mp_GPoint(5, 4))
                oHatchGroup.Children.Add(mp_GPoint(7, 4))

                oHatchGroup.Children.Add(mp_GLine(0, 6, 1, 6))
                oHatchGroup.Children.Add(mp_GLine(3, 8, 4, 8))
                oHatchCtrlGroup.Children.Add(mp_GPoint(0, 0))
                oHatchCtrlGroup.Children.Add(mp_GPoint(7, 7))
                yType = HatchType.Line
            Case HatchStyle.DiagonalBrick
                lWidth = 8
                lHeight = 8
                oHatchGroup.Children.Add(mp_GPoint(0, 7))
                oHatchGroup.Children.Add(mp_GPoint(1, 6))
                oHatchGroup.Children.Add(mp_GPoint(2, 5))
                oHatchGroup.Children.Add(mp_GPoint(3, 4))
                oHatchGroup.Children.Add(mp_GPoint(4, 3))
                oHatchGroup.Children.Add(mp_GPoint(5, 2))
                oHatchGroup.Children.Add(mp_GPoint(6, 1))
                oHatchGroup.Children.Add(mp_GPoint(7, 0))
                oHatchGroup.Children.Add(mp_GPoint(4, 4))
                oHatchGroup.Children.Add(mp_GPoint(5, 5))
                oHatchGroup.Children.Add(mp_GPoint(6, 6))
                oHatchGroup.Children.Add(mp_GPoint(7, 7))
                yType = HatchType.Line
            Case HatchStyle.HorizontalBrick
                lWidth = 8
                lHeight = 8
                oHatchGroup.Children.Add(mp_GPoint(4, 0))
                oHatchGroup.Children.Add(mp_GPoint(4, 1))
                oHatchGroup.Children.Add(mp_GLine(0, 1, 0, 5))
                oHatchGroup.Children.Add(mp_GPoint(4, 6))
                oHatchGroup.Children.Add(mp_GPoint(4, 7))

                oHatchGroup.Children.Add(mp_GLine(1, 2, 7, 2))
                oHatchGroup.Children.Add(mp_GLine(1, 6, 7, 6))
                yType = HatchType.Line
            Case HatchStyle.Weave
                lWidth = 8
                lHeight = 8
                oHatchGroup.Children.Add(mp_GPoint(0, 0))
                oHatchGroup.Children.Add(mp_GPoint(1, 1))
                oHatchGroup.Children.Add(mp_GPoint(2, 2))
                oHatchGroup.Children.Add(mp_GPoint(0, 4))
                oHatchGroup.Children.Add(mp_GPoint(1, 3))
                oHatchGroup.Children.Add(mp_GPoint(3, 1))
                oHatchGroup.Children.Add(mp_GPoint(4, 0))
                oHatchGroup.Children.Add(mp_GPoint(5, 1))
                oHatchGroup.Children.Add(mp_GPoint(6, 2))
                oHatchGroup.Children.Add(mp_GPoint(7, 3))
                oHatchGroup.Children.Add(mp_GPoint(5, 3))
                oHatchGroup.Children.Add(mp_GPoint(4, 4))
                oHatchGroup.Children.Add(mp_GPoint(3, 5))
                oHatchGroup.Children.Add(mp_GPoint(2, 6))
                oHatchGroup.Children.Add(mp_GPoint(1, 7))
                oHatchGroup.Children.Add(mp_GPoint(3, 7))
                oHatchGroup.Children.Add(mp_GPoint(5, 5))
                oHatchGroup.Children.Add(mp_GPoint(6, 6))
                oHatchGroup.Children.Add(mp_GPoint(7, 7))
                yType = HatchType.Line
            Case HatchStyle.Plaid
                lWidth = 8
                lHeight = 8
                oHatchGroup.Children.Add(mp_GLine(0, 0, 3, 0))
                oHatchGroup.Children.Add(mp_GLine(0, 1, 3, 1))

                oHatchGroup.Children.Add(mp_GPoint(0, 2))
                oHatchGroup.Children.Add(mp_GPoint(2, 2))
                oHatchGroup.Children.Add(mp_GPoint(4, 2))
                oHatchGroup.Children.Add(mp_GPoint(6, 2))

                oHatchGroup.Children.Add(mp_GPoint(1, 3))
                oHatchGroup.Children.Add(mp_GPoint(3, 3))
                oHatchGroup.Children.Add(mp_GPoint(5, 3))
                oHatchGroup.Children.Add(mp_GPoint(7, 3))


                oHatchGroup.Children.Add(mp_GPoint(0, 4))
                oHatchGroup.Children.Add(mp_GPoint(2, 4))
                oHatchGroup.Children.Add(mp_GPoint(4, 4))
                oHatchGroup.Children.Add(mp_GPoint(6, 4))

                oHatchGroup.Children.Add(mp_GPoint(1, 5))
                oHatchGroup.Children.Add(mp_GPoint(3, 5))
                oHatchGroup.Children.Add(mp_GPoint(5, 5))
                oHatchGroup.Children.Add(mp_GPoint(7, 5))

                oHatchGroup.Children.Add(mp_GLine(0, 6, 3, 6))
                oHatchGroup.Children.Add(mp_GLine(0, 7, 3, 7))
                oHatchCtrlGroup.Children.Add(mp_GPoint(7, 7))
                yType = HatchType.Line
            Case HatchStyle.Divot
                lWidth = 8
                lHeight = 8
                oHatchGroup.Children.Add(mp_GPoint(0, 1))
                oHatchGroup.Children.Add(mp_GPoint(0, 3))
                oHatchGroup.Children.Add(mp_GPoint(3, 5))
                oHatchGroup.Children.Add(mp_GPoint(3, 7))
                oHatchGroup.Children.Add(mp_GPoint(4, 6))
                oHatchGroup.Children.Add(mp_GPoint(7, 2))
                oHatchCtrlGroup.Children.Add(mp_GPoint(7, 7))
                yType = HatchType.Line
            Case HatchStyle.DottedGrid
                lWidth = 8
                lHeight = 8
                oHatchGroup.Children.Add(mp_GPoint(1, 6))
                oHatchGroup.Children.Add(mp_GPoint(3, 6))
                oHatchGroup.Children.Add(mp_GPoint(5, 6))
                oHatchGroup.Children.Add(mp_GPoint(7, 6))
                oHatchGroup.Children.Add(mp_GPoint(7, 4))
                oHatchGroup.Children.Add(mp_GPoint(7, 2))
                oHatchGroup.Children.Add(mp_GPoint(7, 0))
                oHatchCtrlGroup.Children.Add(mp_GPoint(7, 7))
                yType = HatchType.Line
            Case HatchStyle.DottedDiamond
                lWidth = 8
                lHeight = 8
                oHatchGroup.Children.Add(mp_GPoint(0, 0))
                oHatchGroup.Children.Add(mp_GPoint(2, 2))
                oHatchGroup.Children.Add(mp_GPoint(4, 4))
                oHatchGroup.Children.Add(mp_GPoint(6, 6))
                oHatchGroup.Children.Add(mp_GPoint(2, 6))
                oHatchGroup.Children.Add(mp_GPoint(6, 2))
                yType = HatchType.Line
            Case HatchStyle.Shingle
                lWidth = 8
                lHeight = 8
                oHatchGroup.Children.Add(mp_GPoint(0, 3))
                oHatchGroup.Children.Add(mp_GPoint(1, 4))
                oHatchGroup.Children.Add(mp_GPoint(2, 5))
                oHatchGroup.Children.Add(mp_GPoint(3, 5))
                oHatchGroup.Children.Add(mp_GPoint(4, 6))
                oHatchGroup.Children.Add(mp_GPoint(5, 6))
                oHatchGroup.Children.Add(mp_GPoint(6, 7))
                oHatchGroup.Children.Add(mp_GPoint(4, 4))
                oHatchGroup.Children.Add(mp_GPoint(5, 3))
                oHatchGroup.Children.Add(mp_GPoint(6, 2))
                oHatchGroup.Children.Add(mp_GPoint(7, 2))
                oHatchGroup.Children.Add(mp_GPoint(7, 0))
                oHatchGroup.Children.Add(mp_GPoint(7, 1))
                yType = HatchType.Line
            Case HatchStyle.Trellis
                lWidth = 4
                lHeight = 4
                oHatchGroup.Children.Add(mp_GLine(0, 0, 3, 0))
                oHatchGroup.Children.Add(mp_GLine(1, 1, 2, 1))
                oHatchGroup.Children.Add(mp_GLine(0, 2, 3, 2))
                oHatchGroup.Children.Add(mp_GPoint(0, 3))
                oHatchGroup.Children.Add(mp_GPoint(3, 3))
                yType = HatchType.Line
            Case HatchStyle.Sphere
                lWidth = 8
                lHeight = 8
                oHatchGroup.Children.Add(mp_GLine(1, 0, 3, 0))
                oHatchGroup.Children.Add(mp_GLine(1, 1, 3, 1))
                oHatchGroup.Children.Add(mp_GPoint(0, 2))
                oHatchGroup.Children.Add(mp_GPoint(4, 2))
                oHatchGroup.Children.Add(mp_GLine(1, 3, 2, 3))
                oHatchGroup.Children.Add(mp_GPoint(0, 6))
                oHatchGroup.Children.Add(mp_GPoint(4, 6))
                oHatchGroup.Children.Add(mp_GLine(1, 7, 3, 7))
                oHatchGroup.Children.Add(mp_GLine(5, 7, 6, 7))
                oHatchGroup.Children.Add(mp_GLine(5, 3, 7, 3))
                oHatchGroup.Children.Add(mp_GLine(5, 4, 7, 4))
                oHatchGroup.Children.Add(mp_GLine(5, 5, 7, 5))
                oHatchCtrlGroup.Children.Add(mp_GPoint(7, 7))
                InvertColors(Foreground, Background)
                yType = HatchType.Line
            Case HatchStyle.SmallGrid
                lWidth = 4
                lHeight = 4
                oHatchGroup.Children.Add(mp_GLine(0, 0, 3, 0))
                oHatchGroup.Children.Add(mp_GLine(0, 0, 0, 3))
                yType = HatchType.Line
            Case HatchStyle.SmallCheckerboard
                lWidth = 4
                lHeight = 4
                oHatchGroup.Children.Add(mp_GRect(0, 0, 2, 2))
                oHatchGroup.Children.Add(mp_GRect(2, 2, 2, 2))
                yType = HatchType.Rectangle
            Case HatchStyle.LargeCheckerboard
                lWidth = 8
                lHeight = 8
                oHatchGroup.Children.Add(mp_GRect(0, 0, 4, 4))
                oHatchGroup.Children.Add(mp_GRect(4, 4, 4, 4))
                yType = HatchType.Rectangle
            Case HatchStyle.OutlinedDiamond
                lWidth = 8
                lHeight = 8

                oHatchGroup.Children.Add(mp_GPoint(0, 4))
                oHatchGroup.Children.Add(mp_GPoint(1, 3))
                oHatchGroup.Children.Add(mp_GPoint(2, 2))
                oHatchGroup.Children.Add(mp_GPoint(3, 1))
                oHatchGroup.Children.Add(mp_GPoint(4, 0))

                oHatchGroup.Children.Add(mp_GPoint(5, 1))
                oHatchGroup.Children.Add(mp_GPoint(6, 2))
                oHatchGroup.Children.Add(mp_GPoint(7, 3))

                oHatchGroup.Children.Add(mp_GPoint(7, 5))
                oHatchGroup.Children.Add(mp_GPoint(6, 6))
                oHatchGroup.Children.Add(mp_GPoint(5, 7))

                oHatchGroup.Children.Add(mp_GPoint(3, 7))
                oHatchGroup.Children.Add(mp_GPoint(2, 6))
                oHatchGroup.Children.Add(mp_GPoint(1, 5))
                yType = HatchType.Line
            Case HatchStyle.SolidDiamond
                lWidth = 8
                lHeight = 8

                Dim Lines = {New LineSegment(New Point(0, 4), False),
                                       New LineSegment(New Point(4, 8), False),
                                       New LineSegment(New Point(8, 4), False),
                                       New LineSegment(New Point(4, 0), False)}

                Dim F As New PathFigure(New Point(0, 4), Lines, True)
                Dim P As New PathGeometry({F})
                oHatchGroup.Children.Add(P)
                yType = HatchType.Rectangle
            Case HatchStyle.TwoOverlappingEllipses
                lWidth = 9
                lHeight = 9
                Dim ellipse1 As New EllipseGeometry(New Point(4.5, 4.5), 2, 4.5)
                Dim ellipse2 As New EllipseGeometry(New Point(4.5, 4.5), 4.5, 2)
                oHatchGroup.Children.Add(ellipse1)
                oHatchGroup.Children.Add(ellipse2)
                yType = HatchType.Rectangle
            Case HatchStyle.DotFill
                lWidth = 8.01
                lHeight = 8.01
                oHatchGroup.Children.Add(New EllipseGeometry(New Point(4, 4), 4, 4))
                yType = HatchType.Rectangle
        End Select

        oHatchGroup.Transform = GetShapeTransform(DBrush)

        Dim oBackgroundSquare As New GeometryDrawing(Background, Nothing, New RectangleGeometry(New Rect(0, 0, lWidth, lHeight)))
        Dim oHatchBrush = Foreground
        Dim oHatchPen As New Pen(oHatchBrush, PenThickness)
        Dim oHatchCtrlBrush As New SolidColorBrush(Colors.Red)
        Dim oHatchCtrlPen As New Pen(oHatchCtrlBrush, PenThickness)
        Dim oHatch As GeometryDrawing = Nothing
        Dim oHatchCtrl As GeometryDrawing = Nothing
        Select Case yType
            Case HatchType.Rectangle
                oHatch = New GeometryDrawing(oHatchBrush, Nothing, oHatchGroup)
                If oHatchCtrlGroup.Children.Count > 0 Then
                    oHatchCtrl = New GeometryDrawing(oHatchCtrlBrush, Nothing, oHatchCtrlGroup)
                End If
            Case HatchType.Line
                oHatch = New GeometryDrawing(Nothing, oHatchPen, oHatchGroup)
                If oHatchCtrlGroup.Children.Count > 0 Then
                    oHatchCtrl = New GeometryDrawing(Nothing, oHatchCtrlPen, oHatchCtrlGroup)
                End If
        End Select

        Dim oDrawingGroup As New DrawingGroup
        If bAliased = True Then
            oDrawingGroup.SetValue(RenderOptions.EdgeModeProperty, EdgeMode.Aliased)
        End If
        If Not oHatchCtrl Is Nothing Then
            oDrawingGroup.Children.Add(oHatchCtrl)
        End If
        oDrawingGroup.Children.Add(oBackgroundSquare)
        oDrawingGroup.Children.Add(oHatch)

        DontChange = True
        If DBrush.Drawing Is Nothing Then

            DBrush.Stretch = Stretch.None
            DBrush.ViewportUnits = BrushMappingMode.Absolute
            Dim Rec = New Rect(0, 0, lWidth, lHeight)
            DBrush.Viewport = Rec
            DBrush.ViewboxUnits = BrushMappingMode.Absolute
            DBrush.Viewbox = Rec
            DBrush.TileMode = TileMode.Tile

        End If

        oDrawingGroup.Transform = GetHatchTransform(DBrush)
        DBrush.Drawing = oDrawingGroup
        SetHatchWidth(DBrush, lWidth)
        SetHatchHeight(DBrush, lHeight)
        DontChange = False
    End Sub

    Private Shared Function mp_GLine(X1 As Integer, Y1 As Integer, X2 As Integer, Y2 As Integer) As LineGeometry
        If X1 <> X2 Then
            X2 = X2 + 1
        End If
        If Y1 <> Y2 Then
            Y2 = Y2 + 1
        End If
        Dim oReturn As New LineGeometry(New Point(X1, Y1), New Point(X2, Y2))
        Return oReturn
    End Function

    Private Shared Function mp_GPoint(X1 As Integer, Y1 As Integer) As LineGeometry
        Dim oReturn As New LineGeometry(New Point(X1, Y1), New Point(X1 + 1, Y1 + 1))
        Return oReturn
    End Function

    Private Shared Function mp_GRect(X As Integer, Y As Integer, Width As Integer, Height As Integer) As RectangleGeometry
        Dim oReturn As New RectangleGeometry(New Rect(X, Y, Width, Height))
        Return oReturn
    End Function

    Private Shared Sub InvertColors(ByRef clrForeColor As Brush, ByRef clrBackColor As Brush)
        Dim clrBuff As Brush
        clrBuff = If(clrBackColor IsNot Nothing, clrBackColor.Clone, Nothing)
        clrBackColor = If(clrForeColor IsNot Nothing, clrForeColor.Clone, Nothing)
        clrForeColor = clrBuff
    End Sub

#End Region


End Class
