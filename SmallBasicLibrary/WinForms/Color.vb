Imports System.Windows.Media
Imports Microsoft.SmallVisualBasic.Library

Namespace WinForms
    ''' <summary>
    ''' Contains methods to create and modify colors.
    ''' </summary>
    <SmallVisualBasicType>
    Public NotInheritable Class Color
        Private Shared _random As Random

        ''' <summary>
        ''' Gets a valid random color.
        ''' </summary>
        ''' <returns>
        ''' A valid random color.
        ''' </returns>
        <ReturnValueType(VariableType.Color)>
        Public Shared Function GetRandomColor() As Primitive
            If _random Is Nothing Then
                _random = New Random()
            End If

            Return _colorNames.Keys(_random.Next(_colorNames.Count - 3) + 2)
        End Function

        Private Shared Function InRange(value As Integer, min As Integer, max As Integer) As Integer
            Dim result = System.Math.Min(value, max)
            Return System.Math.Max(result, min)
        End Function

        ''' <summary>
        ''' Creates a color from its red, green, and blue components.
        ''' </summary>
        ''' <param name="red">the 0 to 255 value of the red component</param>
        ''' <param name="green">the 0 to 255 value of the green component</param>
        ''' <param name="blue">the 0 to 255 value of the blue component</param>
        ''' <returns>a string representing the color</returns>
        <ReturnValueType(VariableType.Color)>
        Public Shared Function FromRGB(red As Primitive, green As Primitive, blue As Primitive) As Primitive
            Dim R = InRange(red, 0, 255)
            Dim G = InRange(green, 0, 255)
            Dim B = InRange(blue, 0, 255)
            Return $"#FF{R:X2}{G:X2}{B:X2}"
        End Function

        ''' <summary>
        '''Creates a color from its alpha, red, green, and blue components.
        ''' </summary>
        ''' <param name="alpha">the 0 to 255 value of the color opacity. 0 means the color is fully transparent</param>
        ''' <param name="red">the 0 to 255 value of the red component</param>
        ''' <param name="green">the 0 to 255 value of the green component</param>
        ''' <param name="blue">the 0 to 255 value of the blue component</param>
        ''' <returns>a string representing the color</returns>
        <ReturnValueType(VariableType.Color)>
        Public Shared Function FromARGB(alpha As Primitive, red As Primitive, green As Primitive, blue As Primitive) As Primitive
            Dim A = InRange(alpha, 0, 255)
            Dim R = InRange(red, 0, 255)
            Dim G = InRange(green, 0, 255)
            Dim B = InRange(blue, 0, 255)
            Return $"#{A:X2}{R:X2}{G:X2}{B:X2}"
        End Function

        Friend Shared Function IsNone(color As String) As Boolean
            Return color = "" OrElse color = "0" OrElse color.ToLower() = "none"
        End Function

        ''' <summary>
        ''' Creates a new color based on the given colr, with the transparency set to the given value.
        ''' </summary>
        ''' <param name="color">the color you want to change its transparency</param>
        ''' <param name="percentage">A number between 0 and 100 that represents the percentage of the color transparency</param>
        ''' <returns>a new color with the given transparency</returns>
        <ReturnValueType(VariableType.Color)>
        Public Shared Function ChangeTransparency(color As Primitive, percentage As Primitive) As Primitive
            If IsNone(color) Then Return color

            Dim _color = FromString(color)
            Dim opacity = 1 - InRange(percentage, 0, 100) / 100
            Dim A As Byte = System.Math.Floor(opacity * 255)
            Return $"#{A:X2}{_color.R:X2}{_color.G:X2}{_color.B:X2}"
        End Function

        ''' <summary>
        ''' Gets the transparency percentage of the color
        ''' </summary>
        ''' <param name="color">the color you want to get its transparency percentage</param>
        ''' <returns>a number between 0 and 100 representing the percentage of the color transparency</returns>
        <ReturnValueType(VariableType.Double)>
        Public Shared Function GetTransparency(color As Primitive) As Primitive
            If IsNone(color) Then Return 100
            Dim _color = FromString(color)
            Return System.Math.Round(100 - _color.A * 100 / 255, MidpointRounding.AwayFromZero)
        End Function

        Friend Shared Function FromString(color As String) As System.Windows.Media.Color
            Try
                Return ColorConverter.ConvertFromString(color)
            Catch
                Return System.Windows.Media.Colors.Transparent
            End Try
        End Function

        Friend Shared Function GetPen(penColor As Primitive, penWidth As Primitive) As Pen
            If IsNone(penColor) Then Return Nothing
            Return New Pen(GetBrush(penColor), penWidth)
        End Function

        Friend Shared Function GetBrush(brushColor As Primitive) As Brush
            If IsNone(brushColor) Then Return Nothing
            Return New SolidColorBrush(FromString(brushColor))
        End Function

        ''' <summary>
        ''' Gets the English name of the color if its defined.
        ''' </summary>
        ''' <param name="color">the color you want to get its name</param>
        ''' <returns>the English name of the color</returns>
        <ReturnValueType(VariableType.String)>
        Public Shared Function GetName(color As Primitive) As Primitive
            Return DoGetName(color, True)
        End Function

        ''' <summary>
        ''' Gets the English name of the color if its defined, followed by the transparency percentage of the color.
        ''' </summary>
        ''' <param name="color">the color you want to get its name</param>
        ''' <returns>the English name of the color, followed by the transparency percentage of the color</returns>
        <ReturnValueType(VariableType.String)>
        Public Shared Function GetNameAndTransparency(color As Primitive) As Primitive
            Return DoGetName(color, False)
        End Function

        Private Shared Function DoGetName(color As String, ingnoreTrans As Boolean) As String
            Select Case color.ToLower()
                Case "", "0", "none"
                    Return "None"
                Case Colors.Transparent.ToString().ToLower()
                    Return "Transparent"
            End Select

            Dim _color = FromString(color)
            Dim key = FromRGB(_color.R, _color.G, _color.B)

            Dim name = If(
                _colorNames.ContainsKey(key),
                 _colorNames(key),
                 color
             )

            If ingnoreTrans OrElse _color.A = 255 Then
                Return name
            Else
                Return $"{name} ({GetTransparency(color)}%)"
            End If
        End Function

        ''' <summary>
        ''' Gets the red component of the color
        ''' </summary>
        ''' <param name="color">the input color</param>
        ''' <returns>A number between 0 and 255 that represents the red component of the color</returns>
        <ReturnValueType(VariableType.Double)>
        Public Shared Function GetRedRatio(color As Primitive) As Primitive
            If IsNone(color) Then Return 0
            Dim _color = FromString(color)
            Return _color.R
        End Function

        ''' <summary>
        ''' Gets the green component of the color
        ''' </summary>
        ''' <param name="color">the input color</param>
        ''' <returns>A number between 0 and 255 that represents the green component of the color</returns>
        <ReturnValueType(VariableType.Double)>
        Public Shared Function GetGreenRatio(color As Primitive) As Primitive
            If IsNone(color) Then Return 0
            Dim _color = FromString(color)
            Return _color.G
        End Function

        ''' <summary>
        ''' Gets the alpha component of the color, which is an indicator of the color opacity.
        ''' </summary>
        ''' <param name="color">the input color</param>
        ''' <returns>
        ''' a number between 0 and 255 that represents the alpha component of the color:
        '''   • 0 means a fully transparent color.
        '''   • 255 means a fully opaque solid color.
        ''' </returns>
        <ReturnValueType(VariableType.Double)>
        Public Shared Function GetAlpha(color As Primitive) As Primitive
            If IsNone(color) Then Return 0
            Return FromString(color).A
        End Function

        ''' <summary>
        ''' Gets the blue component of the color
        ''' </summary>
        ''' <param name="color">the input color</param>
        ''' <returns>A number between 0 and 255 that represents the blue component of the color</returns>

        <ReturnValueType(VariableType.Double)>
        Public Shared Function GetBlueRatio(color As Primitive) As Primitive
            If IsNone(color) Then Return 0
            Dim _color = FromString(color)
            Return _color.B
        End Function

        ''' <summary>
        ''' Creates a new color based on the given color, with the red component changed to the given value.
        ''' </summary>
        ''' <param name="color">The input color</param>
        ''' <param name="value">A number between 0 and 255 for the new value of the red component</param>
        ''' <returns>a new color with the red component changed to the given value</returns>
        <ReturnValueType(VariableType.Color)>
        Public Shared Function ChangeRedRatio(color As Primitive, value As Primitive) As Primitive
            If IsNone(color) Then Return color
            Dim _color = FromString(color)
            Return FromARGB(_color.A, value, _color.G, _color.B)
        End Function

        ''' <summary>
        ''' Creates a new color based on the given color, with the green component changed to the given value.
        ''' </summary>
        ''' <param name="color">The input color</param>
        ''' <param name="value">A number between 0 and 255 for the new value of the green component</param>
        ''' <returns>a new color with the green component changed to the given value</returns>
        <ReturnValueType(VariableType.Color)>
        Public Shared Function ChangeGreenRatio(color As Primitive, value As Primitive) As Primitive
            If IsNone(color) Then Return color
            Dim _color = FromString(color)
            Return FromARGB(_color.A, _color.R, value, _color.B)
        End Function


        ''' <summary>
        ''' Creates a new color based on the given color, with the alpha component changed to the given value.
        ''' The alpha component is an indicator of the color opacity, where 0 means a fully transparent color, while 255 means a fully opaque solid color.
        ''' </summary>
        ''' <param name="color">The input color</param>
        ''' <param name="value">A number between 0 and 255 for the new value of the alpha component.</param>
        ''' <returns>a new color with the alpha component changed to the given value</returns>
        <ReturnValueType(VariableType.Color)>
        Public Shared Function ChangeAlpha(color As Primitive, value As Primitive) As Primitive
            If IsNone(color) Then Return color
            Dim _color = FromString(color)
            Return FromARGB(value, _color.R, _color.G, _color.B)
        End Function


        ''' <summary>
        ''' Creates a new color based on the given color, with the blue component changed to the given value.
        ''' </summary>
        ''' <param name="color">The input color</param>
        ''' <param name="value">A number between 0 and 255 for the new value of the blue component</param>
        ''' <returns>a new color with the blue component changed to the given value</returns>
        <ReturnValueType(VariableType.Color)>
        Public Shared Function ChangeBlueRatio(color As Primitive, value As Primitive) As Primitive
            If IsNone(color) Then Return color
            Dim _color = FromString(color)
            Return FromARGB(_color.A, _color.R, _color.G, value)
        End Function


        Private Shared customColors As New List(Of Integer)()
        Private Shared customColorsStr As String

        Shared Sub New()
            customColorsStr = GetSetting("sVB", "Colors", "CustomColors", "")
            If customColorsStr <> "" Then
                customColors.AddRange(
                        From s In customColorsStr.Split(",")
                        Select Convert.ToInt32(s)
                )
            End If
        End Sub

        ''' <summary>
        ''' Shows the color dialog to allow the user to select a color.
        ''' </summary>
        ''' <param name="initialColor">the color that will be selected when the dialog is opened</param>
        ''' <returns>
        ''' If the user selected a color and clicks the OK button, this method will return the selected color.
        ''' If the user canceled the color dialog, this method returms an empty String "".
        ''' </returns>
        <ReturnValueType(VariableType.Color)>
        Public Shared Function ShowDialog(initialColor As Primitive) As Primitive
            Dim c = System.Drawing.Color.Black
            If Not IsNone(initialColor) Then
                Dim wpfColor = FromString(initialColor)
                c = System.Drawing.Color.FromArgb(
                    wpfColor.A,
                    wpfColor.R,
                    wpfColor.G,
                    wpfColor.B
                )
            End If

            Dim dlg As New System.Windows.Forms.ColorDialog() With {
                .AllowFullOpen = True,
                .FullOpen = True,
                .AnyColor = True,
                .Color = c,
                .ShowHelp = True,
                .CustomColors = customColors.ToArray()
            }

            If dlg.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                For Each clr In dlg.CustomColors
                    If Not customColors.Contains(clr) Then
                        customColors.Insert(0, clr)
                    End If
                Next

                c = dlg.Color
                Dim c2 = System.Drawing.Color.FromArgb(0, c.B, c.B, c.R).ToArgb()

                If Not customColors.Contains(c2) Then
                    customColors.Insert(0, c2)
                End If

                If customColors.Count > 16 Then
                    customColors.RemoveAt(16)
                End If

                Dim x = String.Join(",", customColors.ToArray())
                If x <> customColorsStr Then
                    customColorsStr = x
                    SaveSetting("sVB", "Colors", "CustomColors", x)
                End If


                Return FromARGB(c.A, c.R, c.G, c.B)
            End If

            Return ""
        End Function


        Private Shared _colors As Primitive

        ''' <summary>
        ''' Returns an array of all pre-defined colors
        ''' </summary>
        <ReturnValueType(VariableType.Array)>
        Public Shared ReadOnly Property AllColors As Primitive
            Get
                If _colors.IsEmpty Then
                    Dim map = New Dictionary(Of Primitive, Primitive)
                    Dim keys = Color._colorNames.Keys

                    For n = 2 To keys.Count - 1
                        map(n - 1) = keys(n)
                    Next
                    _colors._arrayMap = map
                End If

                Return _colors
            End Get
        End Property

        Public Shared Function GetHexaName(color As System.Windows.Media.Color?) As String
            If color Is Nothing Then Return "None"
            Dim c = color.Value
            Return $"#{c.A:X2}{c.R:X2}{c.G:X2}{c.B:X2}"
        End Function

        Public Shared Function GetHexaName(brush As SolidColorBrush) As String
            If brush Is Nothing Then Return "None"
            Dim color = brush.Color
            Return $"#{color.A:X2}{color.R:X2}{color.G:X2}{color.B:X2}"
        End Function

        Friend Shared _colorNames As New Dictionary(Of String, String) From {
                {"0", "None"},
                {"None", "None"},
                {"#FFF0F8FF", "AliceBlue"},
                {"#FFFAEBD7", "AntiqueWhite"},
                {"#FF7FFFD4", "Aquamarine"},
                {"#FFF0FFFF", "Azure"},
                {"#FFF5F5DC", "Beige"},
                {"#FFFFE4C4", "Bisque"},
                {"#FF000000", "Black"},
                {"#FFFFEBCD", "BlanchedAlmond"},
                {"#FF0000FF", "Blue"},
                {"#FF8A2BE2", "BlueViolet"},
                {"#FFA52A2A", "Brown"},
                {"#FFDEB887", "BurlyWood"},
                {"#FF5F9EA0", "CadetBlue"},
                {"#FF7FFF00", "Chartreuse"},
                {"#FFD2691E", "Chocolate"},
                {"#FFFF7F50", "Coral"},
                {"#FF6495ED", "CornflowerBlue"},
                {"#FFFFF8DC", "Cornsilk"},
                {"#FFDC143C", "Crimson"},
                {"#FF00FFFF", "Cyan"},
                {"#FF00008B", "DarkBlue"},
                {"#FF008B8B", "DarkCyan"},
                {"#FFB8860B", "DarkGoldenrod"},
                {"#FFA9A9A9", "DarkGray"},
                {"#FF006400", "DarkGreen"},
                {"#FFBDB76B", "DarkKhaki"},
                {"#FF8B008B", "DarkMagenta"},
                {"#FF556B2F", "DarkOliveGreen"},
                {"#FFFF8C00", "DarkOrange"},
                {"#FF9932CC", "DarkOrchid"},
                {"#FF8B0000", "DarkRed"},
                {"#FFE9967A", "DarkSalmon"},
                {"#FF8FBC8F", "DarkSeaGreen"},
                {"#FF483D8B", "DarkSlateBlue"},
                {"#FF2F4F4F", "DarkSlateGray"},
                {"#FF00CED1", "DarkTurquoise"},
                {"#FF9400D3", "DarkViolet"},
                {"#FFFF1493", "DeepPink"},
                {"#FF00BFFF", "DeepSkyBlue"},
                {"#FF696969", "DimGray"},
                {"#FF1E90FF", "DodgerBlue"},
                {"#FFB22222", "FireBrick"},
                {"#FFFFFAF0", "FloralWhite"},
                {"#FF228B22", "ForestGreen"},
                {"#FFFF00FF", "Fuchsia"},
                {"#FFDCDCDC", "Gainsboro"},
                {"#FFF8F8FF", "GhostWhite"},
                {"#FFFFD700", "Gold"},
                {"#FFDAA520", "Goldenrod"},
                {"#FF808080", "Gray"},
                {"#FF008000", "Green"},
                {"#FFADFF2F", "GreenYellow"},
                {"#FFF0FFF0", "Honeydew"},
                {"#FFFF69B4", "HotPink"},
                {"#FFCD5C5C", "IndianRed"},
                {"#FF4B0082", "Indigo"},
                {"#FFFFFFF0", "Ivory"},
                {"#FFF0E68C", "Khaki"},
                {"#FFE6E6FA", "Lavender"},
                {"#FFFFF0F5", "LavenderBlush"},
                {"#FF7CFC00", "LawnGreen"},
                {"#FFFFFACD", "LemonChiffon"},
                {"#FFADD8E6", "LightBlue"},
                {"#FFF08080", "LightCoral"},
                {"#FFE0FFFF", "LightCyan"},
                {"#FFFAFAD2", "LightGoldenrodYellow"},
                {"#FFD3D3D3", "LightGray"},
                {"#FF90EE90", "LightGreen"},
                {"#FFFFB6C1", "LightPink"},
                {"#FFFFA07A", "LightSalmon"},
                {"#FF20B2AA", "LightSeaGreen"},
                {"#FF87CEFA", "LightSkyBlue"},
                {"#FF778899", "LightSlateGray"},
                {"#FFB0C4DE", "LightSteelBlue"},
                {"#FFFFFFE0", "LightYellow"},
                {"#FF00FF00", "Lime"},
                {"#FF32CD32", "LimeGreen"},
                {"#FFFAF0E6", "Linen"},
                {"#FF800000", "Maroon"},
                {"#FF66CDAA", "MediumAquamarine"},
                {"#FF0000CD", "MediumBlue"},
                {"#FFBA55D3", "MediumOrchid"},
                {"#FF9370DB", "MediumPurple"},
                {"#FF3CB371", "MediumSeaGreen"},
                {"#FF7B68EE", "MediumSlateBlue"},
                {"#FF00FA9A", "MediumSpringGreen"},
                {"#FF48D1CC", "MediumTurquoise"},
                {"#FFC71585", "MediumVioletRed"},
                {"#FF191970", "MidnightBlue"},
                {"#FFF5FFFA", "MintCream"},
                {"#FFFFE4E1", "MistyRose"},
                {"#FFFFE4B5", "Moccasin"},
                {"#FFFFDEAD", "NavajoWhite"},
                {"#FF000080", "Navy"},
                {"#FFFDF5E6", "OldLace"},
                {"#FF808000", "Olive"},
                {"#FF6B8E23", "OliveDrab"},
                {"#FFFFA500", "Orange"},
                {"#FFFF4500", "OrangeRed"},
                {"#FFDA70D6", "Orchid"},
                {"#FFEEE8AA", "PaleGoldenrod"},
                {"#FF98FB98", "PaleGreen"},
                {"#FFAFEEEE", "PaleTurquoise"},
                {"#FFDB7093", "PaleVioletRed"},
                {"#FFFFEFD5", "PapayaWhip"},
                {"#FFFFDAB9", "PeachPuff"},
                {"#FFCD853F", "Peru"},
                {"#FFFFC0CB", "Pink"},
                {"#FFDDA0DD", "Plum"},
                {"#FFB0E0E6", "PowderBlue"},
                {"#FF800080", "Purple"},
                {"#FFFF0000", "Red"},
                {"#FFBC8F8F", "RosyBrown"},
                {"#FF4169E1", "RoyalBlue"},
                {"#FF8B4513", "SaddleBrown"},
                {"#FFFA8072", "Salmon"},
                {"#FFF4A460", "SandyBrown"},
                {"#FF2E8B57", "SeaGreen"},
                {"#FFFFF5EE", "Seashell"},
                {"#FFA0522D", "Sienna"},
                {"#FFC0C0C0", "Silver"},
                {"#FF87CEEB", "SkyBlue"},
                {"#FF6A5ACD", "SlateBlue"},
                {"#FF708090", "SlateGray"},
                {"#FFFFFAFA", "Snow"},
                {"#FF00FF7F", "SpringGreen"},
                {"#FF4682B4", "SteelBlue"},
                {"#FFD2B48C", "Tan"},
                {"#FF008080", "Teal"},
                {"#FFD8BFD8", "Thistle"},
                {"#FFFF6347", "Tomato"},
                {"#00FFFFFF", "Transparent"},
                {"#FF40E0D0", "Turquoise"},
                {"#FFEE82EE", "Violet"},
                {"#FFF5DEB3", "Wheat"},
                {"#FFFFFFFF", "White"},
                {"#FFF5F5F5", "WhiteSmoke"},
                {"#FFFFFF00", "Yellow"},
                {"#FF9ACD32", "YellowGreen"}
        }


    End Class
End Namespace