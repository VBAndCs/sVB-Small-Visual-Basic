Imports SmallBasicLibrary.Microsoft.SmallBasic.Library

<SmallBasicType>
Public NotInheritable Class Color

    Private Shared Function InRange(value As Integer, min As Integer, max As Integer) As Integer
        Dim result = System.Math.Min(value, max)
        Return System.Math.Max(result, min)
    End Function

    Public Shared Function FromRGB(red As Primitive, green As Primitive, blue As Primitive) As Primitive
        Dim R = InRange(red, 0, 255)
        Dim G = InRange(green, 0, 255)
        Dim B = InRange(green, 0, 255)
        Return $"#{R:X2}{G:X2}{B:X2}"
    End Function

    Public Shared Function FromARGB(alpha As Primitive, red As Primitive, green As Primitive, blue As Primitive) As Primitive
        Dim A = InRange(alpha, 0, 255)
        Dim R = InRange(red, 0, 255)
        Dim G = InRange(green, 0, 255)
        Dim B = InRange(green, 0, 255)
        Return $"#{A:X2}{R:X2}{G:X2}{B:X2}"
    End Function

    Public Shared Function SetTransparency(color As Primitive, percentage As Primitive) As Primitive
        Dim _color = FromString(color)
        Dim A As Byte = System.Math.Round((100 - InRange(percentage, 0, 100)) * 255 / 100)
        Return $"#{A:X2}{_color.R:X2}{_color.G:X2}{_color.B:X2}"
    End Function

    Public Shared Function GetTransparency(color As Primitive) As Primitive
        Dim _color = FromString(color)
        Return System.Math.Round(100 - _color.A * 100 / 255, 1)
    End Function

    Friend Shared Function FromString(color As String) As Media.Color
        Try
            Return CType(ColorConverter.ConvertFromString(color), Media.Color)
        Catch
            Return Colors.Black
        End Try
    End Function

    Public Shared Function GetName(color As Primitive) As Primitive
        Return DoGetName(color, True)
    End Function

    Public Shared Function GetNameAndTransparency(color As Primitive) As Primitive
        Return DoGetName(color, False)
    End Function

    Private Shared Function DoGetName(color As String, ingnoreTrans As Boolean) As String
        If Not color.StartsWith("#") Then Return color
        If color = Transparent.ToString Then Return "Transparent"

        Dim _color = FromString(color)
        Dim key = FromRGB(_color.R, _color.G, _color.B)
        If ColorNames.ContainsKey(key) Then
            If ingnoreTrans OrElse _color.A = 0 Then
                Return ColorNames(key)
            Else
                Return ColorNames(key) & $" ({System.Math.Round(100 - _color.A * 100 / 255)}%)"
            End If
        Else
            Return color
        End If
    End Function

    Public Shared ReadOnly Property AliceBlue As Primitive = "#F0F8FF"
    Public Shared ReadOnly Property AntiqueWhite As Primitive = "#FAEBD7"
    Public Shared ReadOnly Property Aqua As Primitive = "#00FFFF"
    Public Shared ReadOnly Property Aquamarine As Primitive = "#7FFFD4"
    Public Shared ReadOnly Property Azure As Primitive = "#F0FFFF"
    Public Shared ReadOnly Property Beige As Primitive = "#F5F5DC"
    Public Shared ReadOnly Property Bisque As Primitive = "#FFE4C4"
    Public Shared ReadOnly Property Black As Primitive = "#000000"
    Public Shared ReadOnly Property BlanchedAlmond As Primitive = "#FFEBCD"
    Public Shared ReadOnly Property Blue As Primitive = "#0000FF"
    Public Shared ReadOnly Property BlueViolet As Primitive = "#8A2BE2"
    Public Shared ReadOnly Property Brown As Primitive = "#A52A2A"
    Public Shared ReadOnly Property BurlyWood As Primitive = "#DEB887"
    Public Shared ReadOnly Property CadetBlue As Primitive = "#5F9EA0"
    Public Shared ReadOnly Property Chartreuse As Primitive = "#7FFF00"
    Public Shared ReadOnly Property Chocolate As Primitive = "#D2691E"
    Public Shared ReadOnly Property Coral As Primitive = "#FF7F50"
    Public Shared ReadOnly Property CornflowerBlue As Primitive = "#6495ED"
    Public Shared ReadOnly Property Cornsilk As Primitive = "#FFF8DC"
    Public Shared ReadOnly Property Crimson As Primitive = "#DC143C"
    Public Shared ReadOnly Property Cyan As Primitive = "#00FFFF"
    Public Shared ReadOnly Property DarkBlue As Primitive = "#00008B"
    Public Shared ReadOnly Property DarkCyan As Primitive = "#008B8B"
    Public Shared ReadOnly Property DarkGoldenrod As Primitive = "#B8860B"
    Public Shared ReadOnly Property DarkGray As Primitive = "#A9A9A9"
    Public Shared ReadOnly Property DarkGreen As Primitive = "#006400"
    Public Shared ReadOnly Property DarkKhaki As Primitive = "#BDB76B"
    Public Shared ReadOnly Property DarkMagenta As Primitive = "#8B008B"
    Public Shared ReadOnly Property DarkOliveGreen As Primitive = "#556B2F"
    Public Shared ReadOnly Property DarkOrange As Primitive = "#FF8C00"
    Public Shared ReadOnly Property DarkOrchid As Primitive = "#9932CC"
    Public Shared ReadOnly Property DarkRed As Primitive = "#8B0000"
    Public Shared ReadOnly Property DarkSalmon As Primitive = "#E9967A"
    Public Shared ReadOnly Property DarkSeaGreen As Primitive = "#8FBC8F"
    Public Shared ReadOnly Property DarkSlateBlue As Primitive = "#483D8B"
    Public Shared ReadOnly Property DarkSlateGray As Primitive = "#2F4F4F"
    Public Shared ReadOnly Property DarkTurquoise As Primitive = "#00CED1"
    Public Shared ReadOnly Property DarkViolet As Primitive = "#9400D3"
    Public Shared ReadOnly Property DeepPink As Primitive = "#FF1493"
    Public Shared ReadOnly Property DeepSkyBlue As Primitive = "#00BFFF"
    Public Shared ReadOnly Property DimGray As Primitive = "#696969"
    Public Shared ReadOnly Property DodgerBlue As Primitive = "#1E90FF"
    Public Shared ReadOnly Property FireBrick As Primitive = "#B22222"
    Public Shared ReadOnly Property FloralWhite As Primitive = "#FFFAF0"
    Public Shared ReadOnly Property ForestGreen As Primitive = "#228B22"
    Public Shared ReadOnly Property Fuchsia As Primitive = "#FF00FF"
    Public Shared ReadOnly Property Gainsboro As Primitive = "#DCDCDC"
    Public Shared ReadOnly Property GhostWhite As Primitive = "#F8F8FF"
    Public Shared ReadOnly Property Gold As Primitive = "#FFD700"
    Public Shared ReadOnly Property Goldenrod As Primitive = "#DAA520"
    Public Shared ReadOnly Property Gray As Primitive = "#808080"
    Public Shared ReadOnly Property Green As Primitive = "#008000"
    Public Shared ReadOnly Property GreenYellow As Primitive = "#ADFF2F"
    Public Shared ReadOnly Property Honeydew As Primitive = "#F0FFF0"
    Public Shared ReadOnly Property HotPink As Primitive = "#FF69B4"
    Public Shared ReadOnly Property IndianRed As Primitive = "#CD5C5C"
    Public Shared ReadOnly Property Indigo As Primitive = "#4B0082"
    Public Shared ReadOnly Property Ivory As Primitive = "#FFFFF0"
    Public Shared ReadOnly Property Khaki As Primitive = "#F0E68C"
    Public Shared ReadOnly Property Lavender As Primitive = "#E6E6FA"
    Public Shared ReadOnly Property LavenderBlush As Primitive = "#FFF0F5"
    Public Shared ReadOnly Property LawnGreen As Primitive = "#7CFC00"
    Public Shared ReadOnly Property LemonChiffon As Primitive = "#FFFACD"
    Public Shared ReadOnly Property LightBlue As Primitive = "#ADD8E6"
    Public Shared ReadOnly Property LightCoral As Primitive = "#F08080"
    Public Shared ReadOnly Property LightCyan As Primitive = "#E0FFFF"
    Public Shared ReadOnly Property LightGoldenrodYellow As Primitive = "#FAFAD2"
    Public Shared ReadOnly Property LightGray As Primitive = "#D3D3D3"
    Public Shared ReadOnly Property LightGreen As Primitive = "#90EE90"
    Public Shared ReadOnly Property LightPink As Primitive = "#FFB6C1"
    Public Shared ReadOnly Property LightSalmon As Primitive = "#FFA07A"
    Public Shared ReadOnly Property LightSeaGreen As Primitive = "#20B2AA"
    Public Shared ReadOnly Property LightSkyBlue As Primitive = "#87CEFA"
    Public Shared ReadOnly Property LightSlateGray As Primitive = "#778899"
    Public Shared ReadOnly Property LightSteelBlue As Primitive = "#B0C4DE"
    Public Shared ReadOnly Property LightYellow As Primitive = "#FFFFE0"
    Public Shared ReadOnly Property Lime As Primitive = "#00FF00"
    Public Shared ReadOnly Property LimeGreen As Primitive = "#32CD32"
    Public Shared ReadOnly Property Linen As Primitive = "#FAF0E6"
    Public Shared ReadOnly Property Magenta As Primitive = "#FF00FF"
    Public Shared ReadOnly Property Maroon As Primitive = "#800000"
    Public Shared ReadOnly Property MediumAquamarine As Primitive = "#66CDAA"
    Public Shared ReadOnly Property MediumBlue As Primitive = "#0000CD"
    Public Shared ReadOnly Property MediumOrchid As Primitive = "#BA55D3"
    Public Shared ReadOnly Property MediumPurple As Primitive = "#9370DB"
    Public Shared ReadOnly Property MediumSeaGreen As Primitive = "#3CB371"
    Public Shared ReadOnly Property MediumSlateBlue As Primitive = "#7B68EE"
    Public Shared ReadOnly Property MediumSpringGreen As Primitive = "#00FA9A"
    Public Shared ReadOnly Property MediumTurquoise As Primitive = "#48D1CC"
    Public Shared ReadOnly Property MediumVioletRed As Primitive = "#C71585"
    Public Shared ReadOnly Property MidnightBlue As Primitive = "#191970"
    Public Shared ReadOnly Property MintCream As Primitive = "#F5FFFA"
    Public Shared ReadOnly Property MistyRose As Primitive = "#FFE4E1"
    Public Shared ReadOnly Property Moccasin As Primitive = "#FFE4B5"
    Public Shared ReadOnly Property NavajoWhite As Primitive = "#FFDEAD"
    Public Shared ReadOnly Property Navy As Primitive = "#000080"
    Public Shared ReadOnly Property OldLace As Primitive = "#FDF5E6"
    Public Shared ReadOnly Property Olive As Primitive = "#808000"
    Public Shared ReadOnly Property OliveDrab As Primitive = "#6B8E23"
    Public Shared ReadOnly Property Orange As Primitive = "#FFA500"
    Public Shared ReadOnly Property OrangeRed As Primitive = "#FF4500"
    Public Shared ReadOnly Property Orchid As Primitive = "#DA70D6"
    Public Shared ReadOnly Property PaleGoldenrod As Primitive = "#EEE8AA"
    Public Shared ReadOnly Property PaleGreen As Primitive = "#98FB98"
    Public Shared ReadOnly Property PaleTurquoise As Primitive = "#AFEEEE"
    Public Shared ReadOnly Property PaleVioletRed As Primitive = "#DB7093"
    Public Shared ReadOnly Property PapayaWhip As Primitive = "#FFEFD5"
    Public Shared ReadOnly Property PeachPuff As Primitive = "#FFDAB9"
    Public Shared ReadOnly Property Peru As Primitive = "#CD853F"
    Public Shared ReadOnly Property Pink As Primitive = "#FFC0CB"
    Public Shared ReadOnly Property Plum As Primitive = "#DDA0DD"
    Public Shared ReadOnly Property PowderBlue As Primitive = "#B0E0E6"
    Public Shared ReadOnly Property Purple As Primitive = "#800080"
    Public Shared ReadOnly Property Red As Primitive = "#FF0000"
    Public Shared ReadOnly Property RosyBrown As Primitive = "#BC8F8F"
    Public Shared ReadOnly Property RoyalBlue As Primitive = "#4169E1"
    Public Shared ReadOnly Property SaddleBrown As Primitive = "#8B4513"
    Public Shared ReadOnly Property Salmon As Primitive = "#FA8072"
    Public Shared ReadOnly Property SandyBrown As Primitive = "#F4A460"
    Public Shared ReadOnly Property SeaGreen As Primitive = "#2E8B57"
    Public Shared ReadOnly Property Seashell As Primitive = "#FFF5EE"
    Public Shared ReadOnly Property Sienna As Primitive = "#A0522D"
    Public Shared ReadOnly Property Silver As Primitive = "#C0C0C0"
    Public Shared ReadOnly Property SkyBlue As Primitive = "#87CEEB"
    Public Shared ReadOnly Property SlateBlue As Primitive = "#6A5ACD"
    Public Shared ReadOnly Property SlateGray As Primitive = "#708090"
    Public Shared ReadOnly Property Snow As Primitive = "#FFFAFA"
    Public Shared ReadOnly Property SpringGreen As Primitive = "#00FF7F"
    Public Shared ReadOnly Property SteelBlue As Primitive = "#4682B4"
    Public Shared ReadOnly Property Tan As Primitive = "#D2B48C"
    Public Shared ReadOnly Property Teal As Primitive = "#008080"
    Public Shared ReadOnly Property Thistle As Primitive = "#D8BFD8"
    Public Shared ReadOnly Property Tomato As Primitive = "#FF6347"
    Public Shared ReadOnly Property Transparent As Primitive = "#00FFFFFF"
    Public Shared ReadOnly Property Turquoise As Primitive = "#40E0D0"
    Public Shared ReadOnly Property Violet As Primitive = "#EE82EE"
    Public Shared ReadOnly Property Wheat As Primitive = "#F5DEB3"
    Public Shared ReadOnly Property White As Primitive = "#FFFFFF"
    Public Shared ReadOnly Property WhiteSmoke As Primitive = "#F5F5F5"
    Public Shared ReadOnly Property Yellow As Primitive = "#FFFF00"
    Public Shared ReadOnly Property YellowGreen As Primitive = "#9ACD32"

    Private Shared ColorNames As New Dictionary(Of String, String) From {
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
