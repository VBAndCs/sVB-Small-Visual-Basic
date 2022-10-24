
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Media
Imports System.Windows.Media.Imaging
Imports System.Windows.Shapes

Namespace Library
    ''' <summary>
    ''' The Shape object allows you to add, move and rotate shapes to the Graphics window.
    ''' </summary>
    <SmallBasicType>
    Public NotInheritable Class Shapes
        Private Shared _nameGenerationMap As New Dictionary(Of String, Integer)
        Private Shared _rotateTransformMap As New Dictionary(Of String, RotateTransform)
        Private Shared _scaleTransformMap As New Dictionary(Of String, ScaleTransform)
        Private Shared _positionMap As New Dictionary(Of String, Point)

        ''' <summary>
        ''' Adds a rectangle shape with the specified width and height.
        ''' </summary>
        ''' <param name="width">The width of the rectangle shape.</param>
        ''' <param name="height">The height of the rectangle shape.</param>
        ''' <returns>
        ''' The Rectangle shape that was just added to the Graphics Window.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function AddRectangle(width As Primitive, height As Primitive) As Primitive
            Dim name As String = GenerateNewName("Rectangle")
            GraphicsWindow.Invoke(
                Sub()
                    GraphicsWindow.VerifyAccess()
                    Dim shape As New Rectangle With {
                            .Width = width,
                            .Height = height,
                            .Fill = GraphicsWindow._fillBrush,
                            .Stroke = GraphicsWindow._pen.Brush,
                            .StrokeThickness = GraphicsWindow._pen.Thickness
                     }
                    GraphicsWindow.AddShape(name, shape)
                End Sub)
            Return name
        End Function

        ''' <summary>
        ''' Add the geometic path that you create using the GeometricPath type to shapes.
        ''' </summary>
        ''' <param name="penColor">The color used to draw the shape outline</param>
        ''' <param name="penWidth">The width of the shape outline</param>
        ''' <param name="brushColor">The color used to fill the shape</param>
        ''' <returns>The geometic path that was just added to the Graphics Window</returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function AddGeometricPath(
                         penColor As Primitive,
                         penWidth As Primitive,
                         brushColor As Primitive
                    ) As Primitive

            Dim name = GenerateNewName("GeoPath")
            GraphicsWindow.Invoke(
                Sub()
                    GraphicsWindow.VerifyAccess()
                    Dim path = WinForms.GeometricPath._path
                    path.Fill = WinForms.Color.GetBrush(brushColor)
                    path.Stroke = WinForms.Color.GetBrush(penColor)
                    path.StrokeThickness = penWidth
                    GraphicsWindow.AddShape(name, path)
                End Sub)
            Return name
        End Function


        ''' <summary>
        ''' Adds an ellipse shape with the specified width and height.
        ''' </summary>
        ''' <param name="width">
        ''' The width of the ellipse shape.
        ''' </param>
        ''' <param name="height">
        ''' The height of the ellipse shape.
        ''' </param>
        ''' <returns>
        ''' The Ellipse shape that was just added to the Graphics Window.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function AddEllipse(width As Primitive, height As Primitive) As Primitive
            Dim name As String = GenerateNewName("Ellipse")
            GraphicsWindow.Invoke(
                Sub()
                    GraphicsWindow.VerifyAccess()
                    Dim shape As New Ellipse With {
                            .Width = width,
                            .Height = height,
                            .Fill = GraphicsWindow._fillBrush,
                            .Stroke = GraphicsWindow._pen.Brush,
                            .StrokeThickness = GraphicsWindow._pen.Thickness
                     }
                    GraphicsWindow.AddShape(name, shape)
                End Sub)
            Return name
        End Function

        ''' <summary>
        ''' Adds a triangle shape represented by the specified points.
        ''' </summary>
        ''' <param name="x1">
        ''' The x co-ordinate of the first point.
        ''' </param>
        ''' <param name="y1">
        ''' The y co-ordinate of the first point.
        ''' </param>
        ''' <param name="x2">
        ''' The x co-ordinate of the second point.
        ''' </param>
        ''' <param name="y2">
        ''' The y co-ordinate of the second point.
        ''' </param>
        ''' <param name="x3">
        ''' The x co-ordinate of the third point.
        ''' </param>
        ''' <param name="y3">
        ''' The y co-ordinate of the third point.
        ''' </param>
        ''' <returns>
        ''' The Triangle shape that was just added to the Graphics Window.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function AddTriangle(
                         x1 As Primitive, y1 As Primitive,
                         x2 As Primitive, y2 As Primitive,
                         x3 As Primitive, y3 As Primitive
                     ) As Primitive

            Dim name As String = GenerateNewName("Triangle")
            GraphicsWindow.Invoke(
                Sub()
                    GraphicsWindow.AddShape(name, New Polygon With {
                        .Points = New PointCollection() From {
                            New Point(x1, y1),
                            New Point(x2, y2),
                            New Point(x3, y3)
                        },
                        .Fill = GraphicsWindow._fillBrush,
                        .Stroke = GraphicsWindow._pen.Brush,
                        .StrokeThickness = GraphicsWindow._pen.Thickness
                    })
                End Sub)
            Return name
        End Function

        ''' <summary>
        ''' Adds a polygon shape represented by the given points array.       
        ''' </summary>
        ''' <param name="pointsArr">An array of points representing the heads of the polygn. Each item in this array is an array containing the x and y of the point.</param>
        ''' <returns>
        ''' The polygon shape that was just added to the Graphics Window.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function AddPolygon(
                         pointsArr As Primitive
                    ) As Primitive

            If pointsArr.IsEmpty OrElse Not pointsArr.IsArray Then Return ""
            Dim name As String = GenerateNewName("Polygon")

            GraphicsWindow.Invoke(
                Sub()
                    Dim Points As New PointCollection()
                    For Each point In pointsArr._arrayMap.Values
                        Points.Add(New Point(point(1), point(2)))
                    Next

                    GraphicsWindow.AddShape(
                        name,
                        New Polygon With {
                            .Points = Points,
                            .Fill = GraphicsWindow._fillBrush,
                            .Stroke = GraphicsWindow._pen.Brush,
                            .StrokeThickness = GraphicsWindow._pen.Thickness
                        }
                    )
                End Sub)
            Return name
        End Function


        ''' <summary>
        ''' Adds a line between the specified points.
        ''' </summary>
        ''' <param name="x1">
        ''' The x co-ordinate of the first point.
        ''' </param>
        ''' <param name="y1">
        ''' The y co-ordinate of the first point.
        ''' </param>
        ''' <param name="x2">
        ''' The x co-ordinate of the second point.
        ''' </param>
        ''' <param name="y2">
        ''' The y co-ordinate of the second point.
        ''' </param>
        ''' <returns>
        ''' The line that was just added to the Graphics Window.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function AddLine(
                          x1 As Primitive, y1 As Primitive,
                          x2 As Primitive, y2 As Primitive
                    ) As Primitive

            Dim name As String = GenerateNewName("Line")
            GraphicsWindow.Invoke(
                Sub()
                    GraphicsWindow.AddShape(name, New Line With {
                        .X1 = x1,
                        .Y1 = y1,
                        .X2 = x2,
                        .Y2 = y2,
                        .Stroke = GraphicsWindow._pen.Brush,
                        .StrokeThickness = GraphicsWindow._pen.Thickness
                    })
                End Sub)
            Return name
        End Function

        ''' <summary>
        ''' Adds an image as a shape that can be moved, animated or rotated.
        ''' </summary>
        ''' <param name="imageName">
        ''' The name of the image to draw.
        ''' </param>
        ''' <returns>
        ''' The image that was just added to the Graphics Window.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function AddImage(imageName As Primitive) As Primitive
            Dim name As String = GenerateNewName("Image")
            GraphicsWindow.Invoke(
                Sub()
                    Dim image1 As New Image
                    Dim bitmap As BitmapSource = ImageList.GetBitmap(imageName)
                    If bitmap IsNot Nothing Then
                        image1.Source = bitmap
                        If bitmap.PixelWidth <> 0 AndAlso bitmap.PixelHeight <> 0 Then
                            image1.Width = bitmap.PixelWidth
                            image1.Height = bitmap.PixelHeight
                        End If
                        GraphicsWindow.AddShape(name, image1)
                    End If
                End Sub)
            Return name
        End Function

        ''' <summary>
        ''' Adds some text as a shape that can be moved, animated or rotated.
        ''' </summary>
        ''' <param name="text">
        ''' The text to add.
        ''' </param>
        ''' <returns>
        ''' The text shape that was just added to the Graphics Window.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function AddText(text As Primitive) As Primitive
            Dim name As String = GenerateNewName("Text")
            GraphicsWindow.Invoke(
                Sub()
                    GraphicsWindow.VerifyAccess()
                    Dim shape As New TextBlock With {
                        .Text = text,
                        .Foreground = GraphicsWindow._fillBrush,
                        .FontFamily = GraphicsWindow._fontFamily,
                        .FontStyle = GraphicsWindow._fontStyle,
                        .FontSize = GraphicsWindow._fontSize,
                        .FontWeight = GraphicsWindow._fontWeight
                     }
                    GraphicsWindow.AddShape(name, shape)
                End Sub)
            Return name
        End Function

        ''' <summary>
        ''' Sets the text of a text shape. 
        ''' </summary>
        ''' <param name="shapeName">
        ''' The name of the text shape.
        ''' </param>
        ''' <param name="text">
        ''' The new text value to set.
        ''' </param>
        Public Shared Sub SetText(shapeName As Primitive, text As Primitive)
            Dim value As UIElement = Nothing
            If Not GraphicsWindow._objectsMap.TryGetValue(shapeName, value) Then
                Return
            End If
            Dim textBlock1 As TextBlock = TryCast(value, TextBlock)
            If textBlock1 IsNot Nothing Then
                GraphicsWindow.BeginInvoke(Sub() textBlock1.Text = text)
            End If
        End Sub

        ''' <summary>
        ''' Removes a shape from the Graphics Window.
        ''' </summary>
        ''' <param name="shapeName">
        ''' The name of the shape that needs to be removed.
        ''' </param>
        Public Shared Sub Remove(shapeName As Primitive)
            GraphicsWindow.RemoveShape(shapeName)
        End Sub

        ''' <summary>
        ''' Moves the shape with the specified name to a new position.
        ''' </summary>
        ''' <param name="shapeName">
        ''' The name of the shape to move.
        ''' </param>
        ''' <param name="x">
        ''' The x co-ordinate of the new position.
        ''' </param>
        ''' <param name="y">
        ''' The y co-ordinate of the new position.
        ''' </param>
        Public Shared Sub Move(shapeName As Primitive, x As Primitive, y As Primitive)
            Dim obj As UIElement = Nothing
            If GraphicsWindow._objectsMap.TryGetValue(shapeName, obj) Then
                _positionMap(shapeName) = New Point(x, y)
                GraphicsWindow.BeginInvoke(
                    Sub()
                        obj.BeginAnimation(Canvas.LeftProperty, Nothing)
                        obj.BeginAnimation(Canvas.TopProperty, Nothing)
                        Canvas.SetLeft(obj, x)
                        Canvas.SetTop(obj, y)
                    End Sub)
            End If
        End Sub

        ''' <summary>
        ''' Rotates the shape with the specified name to the specified angle.
        ''' </summary>
        ''' <param name="shapeName">
        ''' The name of the shape to rotate.
        ''' </param>
        ''' <param name="angle">
        ''' The angle to rotate the shape.
        ''' </param>
        Public Shared Sub Rotate(shapeName As Primitive, angle As Primitive)
            Dim obj As UIElement = Nothing
            If Not GraphicsWindow._objectsMap.TryGetValue(shapeName, obj) Then
                Return
            End If
            GraphicsWindow.BeginInvoke(
                Sub()
                    If Not (TypeOf obj.RenderTransform Is TransformGroup) Then
                        obj.RenderTransform = New TransformGroup
                    End If

                    Dim value As System.Windows.Media.RotateTransform = Nothing
                    If Not _rotateTransformMap.TryGetValue(shapeName, value) Then
                        value = New RotateTransform
                        _rotateTransformMap(shapeName) = value
                        Dim frameworkElement1 As FrameworkElement = TryCast(obj, FrameworkElement)
                        If frameworkElement1 IsNot Nothing Then
                            value.CenterX = frameworkElement1.ActualWidth / 2.0
                            value.CenterY = frameworkElement1.ActualHeight / 2.0
                        End If
                        CType(obj.RenderTransform, TransformGroup).Children.Add(value)
                    End If
                    value.Angle = angle
                End Sub)
        End Sub

        ''' <summary>
        ''' Scales the shape using the specified zoom levels.  Minimum is 0.1 and maximum is 20.
        ''' </summary>
        ''' <param name="shapeName">
        ''' The name of the shape to zoom.
        ''' </param>
        ''' <param name="scaleX">
        ''' The x-axis zoom level.
        ''' </param>
        ''' <param name="scaleY">
        ''' The y-axis zoom level.
        ''' </param>
        Public Shared Sub Zoom(shapeName As Primitive, scaleX As Primitive, scaleY As Primitive)
            Dim obj As UIElement = Nothing
            If Not GraphicsWindow._objectsMap.TryGetValue(shapeName, obj) Then
                Return
            End If
            scaleX = Math.Min(Math.Max(scaleX, 0.1), 20.0)
            scaleY = Math.Min(Math.Max(scaleY, 0.1), 20.0)
            GraphicsWindow.BeginInvoke(
                Sub()
                    If Not (TypeOf obj.RenderTransform Is TransformGroup) Then
                        obj.RenderTransform = New TransformGroup
                    End If

                    Dim value As System.Windows.Media.ScaleTransform = Nothing
                    If Not _scaleTransformMap.TryGetValue(shapeName, value) Then
                        value = New ScaleTransform
                        _scaleTransformMap(shapeName) = value
                        Dim frameworkElement1 As FrameworkElement = TryCast(obj, FrameworkElement)
                        If frameworkElement1 IsNot Nothing Then
                            value.CenterX = frameworkElement1.ActualWidth / 2.0
                            value.CenterY = frameworkElement1.ActualHeight / 2.0
                        End If
                        CType(obj.RenderTransform, TransformGroup).Children.Add(value)
                    End If
                    value.ScaleX = scaleX
                    value.ScaleY = scaleY
                End Sub)
        End Sub

        ''' <summary>
        ''' Animates a shape with the specified name to a new position.
        ''' </summary>
        ''' <param name="shapeName">
        ''' The name of the shape to move.
        ''' </param>
        ''' <param name="x">
        ''' The x co-ordinate of the new position.
        ''' </param>
        ''' <param name="y">
        ''' The y co-ordinate of the new position.
        ''' </param>
        ''' <param name="duration">
        ''' The time for the animation, in milliseconds.
        ''' </param>
        Public Shared Sub Animate(shapeName As Primitive, x As Primitive, y As Primitive, duration As Primitive)
            Dim obj As UIElement = Nothing
            If GraphicsWindow._objectsMap.TryGetValue(shapeName, obj) Then
                _positionMap(shapeName) = New Point(x, y)
                GraphicsWindow.Invoke(
                    Sub()
                        GraphicsWindow.DoubleAnimateProperty(obj, Canvas.LeftProperty, x, duration)
                        GraphicsWindow.DoubleAnimateProperty(obj, Canvas.TopProperty, y, duration)
                    End Sub)
            End If
        End Sub

        ''' <summary>
        ''' Gets the left co-ordinate of the specified shape.
        ''' </summary>
        ''' <param name="shapeName">
        ''' The name of the shape.
        ''' </param>
        ''' <returns>
        ''' The left co-ordinate of the shape.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Function GetLeft(shapeName As Primitive) As Primitive
            Dim __ As UIElement = Nothing
            If Not GraphicsWindow._objectsMap.TryGetValue(shapeName, __) Then
                Return 0
            End If

            Dim value2 As System.Windows.Point = Nothing
            If _positionMap.TryGetValue(shapeName, value2) Then
                Return value2.X
            End If

            Return 0
        End Function

        ''' <summary>
        ''' Gets the top co-ordinate of the specified shape.
        ''' </summary>
        ''' <param name="shapeName">
        ''' The name of the shape.
        ''' </param>
        ''' <returns>
        ''' The top co-ordinate of the shape.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Function GetTop(shapeName As Primitive) As Primitive
            Dim __ As UIElement = Nothing
            If Not GraphicsWindow._objectsMap.TryGetValue(shapeName, __) Then
                Return 0
            End If

            Dim value2 As Point = Nothing
            If _positionMap.TryGetValue(shapeName, value2) Then
                Return value2.Y
            End If

            Return 0
        End Function

        ''' <summary>
        ''' Gets the opacity of a shape.
        ''' </summary>
        ''' <param name="shapeName">
        ''' The name of the shape.
        ''' </param>
        ''' <returns>
        ''' The opacity of the object as a value between 0 and 100.  0 is completely transparent and 100 is completely opaque.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Function GetOpacity(shapeName As Primitive) As Primitive
            Dim obj As UIElement = Nothing
            If Not GraphicsWindow._objectsMap.TryGetValue(shapeName, obj) Then
                Return 0
            End If

            Return CType(GraphicsWindow.InvokeWithReturn(Function() CType(obj.Opacity, Primitive) * CType(100, Primitive)), Primitive)
        End Function

        ''' <summary>
        ''' Sets how opaque a shape should render.
        ''' </summary>
        ''' <param name="shapeName">
        ''' The name of the shape.
        ''' </param>
        ''' <param name="level">
        ''' The opacity level ranging from 0 to 100.  0 is completely transparent and 100 is completely opaque.
        ''' </param>
        Public Shared Sub SetOpacity(shapeName As Primitive, level As Primitive)
            Dim obj As UIElement = Nothing
            If GraphicsWindow._objectsMap.TryGetValue(shapeName, obj) Then
                GraphicsWindow.Invoke(Sub() obj.Opacity = Math.Min(100, Math.Max(0, level)) / 100)
            End If
        End Sub

        ''' <summary>
        ''' Hides an already added shape.
        ''' </summary>
        ''' <param name="shapeName">
        ''' The name of the shape.
        ''' </param>
        Public Shared Sub HideShape(shapeName As Primitive)
            Dim obj As UIElement = Nothing
            If GraphicsWindow._objectsMap.TryGetValue(shapeName, obj) Then
                GraphicsWindow.Invoke(Sub() obj.Visibility = Visibility.Collapsed)
            End If
        End Sub

        ''' <summary>
        ''' Shows a previously hidden shape.
        ''' </summary>
        ''' <param name="shapeName">
        ''' The name of the shape.
        ''' </param>
        Public Shared Sub ShowShape(shapeName As Primitive)
            Dim obj As UIElement = Nothing
            If GraphicsWindow._objectsMap.TryGetValue(shapeName, obj) Then
                GraphicsWindow.Invoke(Sub() obj.Visibility = Visibility.Visible)
            End If
        End Sub

        Friend Shared Function GenerateNewName(prefix As String) As String
            Dim value As Integer = 0
            _nameGenerationMap.TryGetValue(prefix, value)
            value += 1
            _nameGenerationMap(prefix) = value
            Return prefix & value
        End Function
    End Class
End Namespace
