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
                _random = New Random(Now.Ticks Mod Integer.MaxValue)
            End If

            Return $"#{_random.Next(256):X2}{_random.Next(256):X2}{_random.Next(256):X2}"
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
            Return $"#{R:X2}{G:X2}{B:X2}"
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
        ''' Crates a new color based on the given colr, with the transparency set to the given value.
        ''' </summary>
        ''' <param name="color">the color you want to change its transparency</param>
        ''' <param name="percentage">a 0 to 100 value that represents the percentage of the transparency of the color</param>
        ''' <returns>a new color with the given transparency</returns>
        Public Shared Function ChangeTransparency(color As Primitive, percentage As Primitive) As Primitive
            If IsNone(color) Then Return color

            Dim _color = FromString(color)
            Dim A As Byte = System.Math.Round((100 - InRange(percentage, 0, 100)) * 255 / 100)
            Return $"#{A:X2}{_color.R:X2}{_color.G:X2}{_color.B:X2}"
        End Function

        ''' <summary>
        ''' reterns the transparency percentage of the color
        ''' </summary>
        ''' <param name="color">the color you want to get its transparency percentage</param>
        ''' <returns>a 0 to 100 value represents the transparency percentage of the color</returns>
        <ReturnValueType(VariableType.Double)>
        Public Shared Function GetTransparency(color As Primitive) As Primitive
            If IsNone(color) Then Return 100
            Dim _color = FromString(color)
            Return System.Math.Round(100 - _color.A * 100 / 255, 1)
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
        ''' Returms the English name name of the color if its defined.
        ''' </summary>
        ''' <param name="color">the color you want to get its name</param>
        ''' <returns>the English name name of the color</returns>
        <ReturnValueType(VariableType.String)>
        Public Shared Function GetName(color As Primitive) As Primitive
            Return DoGetName(color, True)
        End Function

        ''' <summary>
        ''' Returms the English name name of the color if its defined, followed by the transparency percentage of the color.
        ''' </summary>
        ''' <param name="color">the color you want to get its name</param>
        ''' <returns>the English name name of the color, followed by the transparency percentage of the color</returns>
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

                Case Else
                    If Not color.StartsWith("#") Then Return color
            End Select

            Dim _color = FromString(color)
            Dim key = FromRGB(_color.R, _color.G, _color.B)

            If _colorNames.ContainsKey(key) Then
                If ingnoreTrans OrElse _color.A = 255 Then
                    Return _colorNames(key)
                Else
                    Return _colorNames(key) & $" ({System.Math.Round(100 - _color.A * 100 / 255)}%)"
                End If
            Else
                Return color
            End If
        End Function

        ''' <summary>
        ''' Gets the red component of the color
        ''' </summary>
        ''' <param name="color">the input color</param>
        ''' <returns>A number betwwen 0 and 255 thet represents the red ration in the color</returns>
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
        ''' <returns>A number betwwen 0 and 255 thet represents the green ration in the color</returns>
        <ReturnValueType(VariableType.Double)>
        Public Shared Function GetGreenRatio(color As Primitive) As Primitive
            If IsNone(color) Then Return 0
            Dim _color = FromString(color)
            Return _color.G
        End Function

        ''' <summary>
        ''' Gets the blue component of the color
        ''' </summary>
        ''' <param name="color">the input color</param>
        ''' <returns>A number betwwen 0 and 255 thet represents the blue ration in the color</returns>

        <ReturnValueType(VariableType.Double)>
        Public Shared Function GetBlueRatio(color As Primitive) As Primitive
            If IsNone(color) Then Return 0
            Dim _color = FromString(color)
            Return _color.B
        End Function

        ''' <summary>
        ''' Creates a new color based on the given color, with the red component changed to the given value.
        ''' </summary>
        ''' <param name="color">the input color</param>
        ''' <param name="value">the new value of the red component</param>
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
        ''' <param name="color">the input color</param>
        ''' <param name="value">the new value of the green component</param>
        ''' <returns>a new color with the green component changed to the given value</returns>
        <ReturnValueType(VariableType.Color)>
        Public Shared Function ChangeGreenRatio(color As Primitive, value As Primitive) As Primitive
            If IsNone(color) Then Return color
            Dim _color = FromString(color)
            Return FromARGB(_color.A, _color.R, value, _color.B)
        End Function

        ''' <summary>
        ''' Creates a new color based on the given color, with the blue component changed to the given value.
        ''' </summary>
        ''' <param name="color">the input color</param>
        ''' <param name="value">the new value of the blue component</param>
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
        ''' <returns>a new color with the blue component changed to the given value</returns>
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
                    Dim num = 1

                    For Each key In Color._colorNames.Keys
                        map(num) = key
                        num += 1
                    Next
                    _colors._arrayMap = map
                End If
                Return _colors
            End Get
        End Property


        Friend Shared _colorNames As New Dictionary(Of String, String) From {
                {"0", "None"},
                {"None", "None"},
                {"#F0F8FF", "AliceBlue"},
                {"#FAEBD7", "AntiqueWhite"},
                {"#7FFFD4", "Aquamarine"},
                {"#F0FFFF", "Azure"},
                {"#F5F5DC", "Beige"},
                {"#FFE4C4", "Bisque"},
                {"#000000", "Black"},
                {"#FFEBCD", "BlanchedAlmond"},
                {"#0000FF", "Blue"},
                {"#8A2BE2", "BlueViolet"},
                {"#A52A2A", "Brown"},
                {"#DEB887", "BurlyWood"},
                {"#5F9EA0", "CadetBlue"},
                {"#7FFF00", "Chartreuse"},
                {"#D2691E", "Chocolate"},
                {"#FF7F50", "Coral"},
                {"#6495ED", "CornflowerBlue"},
                {"#FFF8DC", "Cornsilk"},
                {"#DC143C", "Crimson"},
                {"#00FFFF", "Cyan"},
                {"#00008B", "DarkBlue"},
                {"#008B8B", "DarkCyan"},
                {"#B8860B", "DarkGoldenrod"},
                {"#A9A9A9", "DarkGray"},
                {"#006400", "DarkGreen"},
                {"#BDB76B", "DarkKhaki"},
                {"#8B008B", "DarkMagenta"},
                {"#556B2F", "DarkOliveGreen"},
                {"#FF8C00", "DarkOrange"},
                {"#9932CC", "DarkOrchid"},
                {"#8B0000", "DarkRed"},
                {"#E9967A", "DarkSalmon"},
                {"#8FBC8F", "DarkSeaGreen"},
                {"#483D8B", "DarkSlateBlue"},
                {"#2F4F4F", "DarkSlateGray"},
                {"#00CED1", "DarkTurquoise"},
                {"#9400D3", "DarkViolet"},
                {"#FF1493", "DeepPink"},
                {"#00BFFF", "DeepSkyBlue"},
                {"#696969", "DimGray"},
                {"#1E90FF", "DodgerBlue"},
                {"#B22222", "FireBrick"},
                {"#FFFAF0", "FloralWhite"},
                {"#228B22", "ForestGreen"},
                {"#FF00FF", "Fuchsia"},
                {"#DCDCDC", "Gainsboro"},
                {"#F8F8FF", "GhostWhite"},
                {"#FFD700", "Gold"},
                {"#DAA520", "Goldenrod"},
                {"#808080", "Gray"},
                {"#008000", "Green"},
                {"#ADFF2F", "GreenYellow"},
                {"#F0FFF0", "Honeydew"},
                {"#FF69B4", "HotPink"},
                {"#CD5C5C", "IndianRed"},
                {"#4B0082", "Indigo"},
                {"#FFFFF0", "Ivory"},
                {"#F0E68C", "Khaki"},
                {"#E6E6FA", "Lavender"},
                {"#FFF0F5", "LavenderBlush"},
                {"#7CFC00", "LawnGreen"},
                {"#FFFACD", "LemonChiffon"},
                {"#ADD8E6", "LightBlue"},
                {"#F08080", "LightCoral"},
                {"#E0FFFF", "LightCyan"},
                {"#FAFAD2", "LightGoldenrodYellow"},
                {"#D3D3D3", "LightGray"},
                {"#90EE90", "LightGreen"},
                {"#FFB6C1", "LightPink"},
                {"#FFA07A", "LightSalmon"},
                {"#20B2AA", "LightSeaGreen"},
                {"#87CEFA", "LightSkyBlue"},
                {"#778899", "LightSlateGray"},
                {"#B0C4DE", "LightSteelBlue"},
                {"#FFFFE0", "LightYellow"},
                {"#00FF00", "Lime"},
                {"#32CD32", "LimeGreen"},
                {"#FAF0E6", "Linen"},
                {"#800000", "Maroon"},
                {"#66CDAA", "MediumAquamarine"},
                {"#0000CD", "MediumBlue"},
                {"#BA55D3", "MediumOrchid"},
                {"#9370DB", "MediumPurple"},
                {"#3CB371", "MediumSeaGreen"},
                {"#7B68EE", "MediumSlateBlue"},
                {"#00FA9A", "MediumSpringGreen"},
                {"#48D1CC", "MediumTurquoise"},
                {"#C71585", "MediumVioletRed"},
                {"#191970", "MidnightBlue"},
                {"#F5FFFA", "MintCream"},
                {"#FFE4E1", "MistyRose"},
                {"#FFE4B5", "Moccasin"},
                {"#FFDEAD", "NavajoWhite"},
                {"#000080", "Navy"},
                {"#FDF5E6", "OldLace"},
                {"#808000", "Olive"},
                {"#6B8E23", "OliveDrab"},
                {"#FFA500", "Orange"},
                {"#FF4500", "OrangeRed"},
                {"#DA70D6", "Orchid"},
                {"#EEE8AA", "PaleGoldenrod"},
                {"#98FB98", "PaleGreen"},
                {"#AFEEEE", "PaleTurquoise"},
                {"#DB7093", "PaleVioletRed"},
                {"#FFEFD5", "PapayaWhip"},
                {"#FFDAB9", "PeachPuff"},
                {"#CD853F", "Peru"},
                {"#FFC0CB", "Pink"},
                {"#DDA0DD", "Plum"},
                {"#B0E0E6", "PowderBlue"},
                {"#800080", "Purple"},
                {"#FF0000", "Red"},
                {"#BC8F8F", "RosyBrown"},
                {"#4169E1", "RoyalBlue"},
                {"#8B4513", "SaddleBrown"},
                {"#FA8072", "Salmon"},
                {"#F4A460", "SandyBrown"},
                {"#2E8B57", "SeaGreen"},
                {"#FFF5EE", "Seashell"},
                {"#A0522D", "Sienna"},
                {"#C0C0C0", "Silver"},
                {"#87CEEB", "SkyBlue"},
                {"#6A5ACD", "SlateBlue"},
                {"#708090", "SlateGray"},
                {"#FFFAFA", "Snow"},
                {"#00FF7F", "SpringGreen"},
                {"#4682B4", "SteelBlue"},
                {"#D2B48C", "Tan"},
                {"#008080", "Teal"},
                {"#D8BFD8", "Thistle"},
                {"#FF6347", "Tomato"},
                {"#00FFFFFF", "Transparent"},
                {"#40E0D0", "Turquoise"},
                {"#EE82EE", "Violet"},
                {"#F5DEB3", "Wheat"},
                {"#FFFFFF", "White"},
                {"#F5F5F5", "WhiteSmoke"},
                {"#FFFF00", "Yellow"},
                {"#9ACD32", "YellowGreen"}
        }

    End Class
End Namespace