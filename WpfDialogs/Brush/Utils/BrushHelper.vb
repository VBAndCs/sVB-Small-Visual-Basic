Public Class BrushHelper

    Public Shared Function BrushesAreEqual(ByVal aBrush1 As Brush, ByVal aBrush2 As Brush) As Boolean
        If aBrush1 Is Nothing Then
            If aBrush2 Is Nothing Then Return True
            Return False
        ElseIf aBrush2 Is Nothing Then
            Return False
        End If

        If aBrush1.GetType() IsNot aBrush2.GetType() Then Return False

        If Not TransformsAreEqual(aBrush1.Transform, aBrush2.Transform) Then Return False
        If Not TransformsAreEqual(aBrush1.RelativeTransform, aBrush2.RelativeTransform) Then Return False

        If TypeOf aBrush1 Is SolidColorBrush Then
            Dim SolidBrush1 = CType(aBrush1, SolidColorBrush)
            Dim SolidBrush2 = CType(aBrush2, SolidColorBrush)
            Return SolidBrush1.Color = SolidBrush2.Color AndAlso
                        SolidBrush1.Opacity = SolidBrush2.Opacity

        ElseIf TypeOf aBrush1 Is LinearGradientBrush Then
            Dim LinearBrush1 = CType(aBrush1, LinearGradientBrush)
            Dim LinearBrush2 = CType(aBrush2, LinearGradientBrush)
            Return LinearBrush1.ColorInterpolationMode = LinearBrush2.ColorInterpolationMode AndAlso
                        LinearBrush1.StartPoint = LinearBrush2.StartPoint AndAlso
                        LinearBrush1.EndPoint = LinearBrush2.EndPoint AndAlso
                        LinearBrush1.SpreadMethod = LinearBrush2.SpreadMethod AndAlso
                        LinearBrush1.MappingMode = LinearBrush2.MappingMode AndAlso
                        LinearBrush1.Opacity = LinearBrush2.Opacity AndAlso
                        GradientStopsAreEqual(LinearBrush1.GradientStops, LinearBrush2.GradientStops)

        ElseIf TypeOf aBrush1 Is RadialGradientBrush Then
            Dim RadialBrush1 = CType(aBrush1, RadialGradientBrush)
            Dim RadialBrush2 = CType(aBrush2, RadialGradientBrush)
            Return RadialBrush1.ColorInterpolationMode = RadialBrush2.ColorInterpolationMode AndAlso
                        RadialBrush1.GradientOrigin = RadialBrush2.GradientOrigin AndAlso
                        RadialBrush1.SpreadMethod = RadialBrush2.SpreadMethod AndAlso
                        RadialBrush1.MappingMode = RadialBrush2.MappingMode AndAlso
                        RadialBrush1.Opacity = RadialBrush2.Opacity AndAlso
                        RadialBrush1.RadiusX = RadialBrush2.RadiusX AndAlso
                        RadialBrush1.RadiusY = RadialBrush2.RadiusY AndAlso
                        GradientStopsAreEqual(RadialBrush1.GradientStops, RadialBrush2.GradientStops)

        ElseIf TypeOf aBrush1 Is TileBrush Then
            Dim TileBrush1 = CType(aBrush1, TileBrush)
            Dim TileBrush2 = CType(aBrush2, TileBrush)
            Dim Result = TileBrush1.AlignmentY = TileBrush2.AlignmentY AndAlso
                                  TileBrush1.Opacity = TileBrush2.Opacity AndAlso
                                  TileBrush1.Stretch = TileBrush2.Stretch AndAlso
                                  TileBrush1.TileMode = TileBrush2.TileMode AndAlso
                                  TileBrush1.Viewbox = TileBrush2.Viewbox AndAlso
                                  TileBrush1.ViewboxUnits = TileBrush2.ViewboxUnits AndAlso
                                  TileBrush1.Viewport = TileBrush2.Viewport AndAlso
                                  TileBrush1.ViewportUnits = TileBrush2.ViewportUnits

            If TypeOf aBrush1 Is ImageBrush Then
                Dim ImageBrush1 = CType(aBrush1, ImageBrush)
                Dim ImageBrush2 = CType(aBrush2, ImageBrush)
                Return ImageBrush1.ImageSource.ToString = ImageBrush2.ImageSource.ToString AndAlso Result

            ElseIf TypeOf aBrush1 Is DrawingBrush Then
                Return HatchBrushes.GetHatchStyle(aBrush1) = HatchBrushes.GetHatchStyle(aBrush2) AndAlso
                            HatchBrushes.GetPenThickness(aBrush1) = HatchBrushes.GetPenThickness(aBrush2) AndAlso
                            BrushesAreEqual(HatchBrushes.GetBackground(aBrush1), HatchBrushes.GetBackground(aBrush2)) AndAlso
                            BrushesAreEqual(HatchBrushes.GetForeground(aBrush1), HatchBrushes.GetForeground(aBrush2)) AndAlso
                            TransformsAreEqual(HatchBrushes.GetHatchTransform(aBrush1), HatchBrushes.GetHatchTransform(aBrush2)) AndAlso
                            TransformsAreEqual(HatchBrushes.GetShapeTransform(aBrush1), HatchBrushes.GetShapeTransform(aBrush2))
            End If
        End If
        Return False
    End Function

    Friend Shared Function GradientStopsAreEqual(GradientStops1 As GradientStopCollection, GradientStops2 As GradientStopCollection) As Boolean
        If GradientStops1.Count <> GradientStops2.Count Then Return False
        For i As Integer = 0 To GradientStops1.Count - 1
            If GradientStops1(i).Color <> GradientStops2(i).Color Then Return False
            If GradientStops1(i).Offset <> GradientStops2(i).Offset Then Return False
        Next i
        Return True
    End Function

    Public Shared Function TransformsAreEqual(Transform1 As Transform, Transform2 As Transform) As Boolean
        If Transform1 Is Nothing Then
            If Transform2 Is Nothing OrElse Transform2.Value = Matrix.Identity Then Return True
            Return False
        ElseIf Transform2 Is Nothing Then
            If Transform1.Value = Matrix.Identity Then Return True
            Return False
        End If

        Return (Transform1.Value = Transform2.Value)

    End Function



End Class
