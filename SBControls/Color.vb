Imports Microsoft.SmallBasic.Library

<SmallBasicType>
Public Module Color
    Public Function FromRGB(red As Primitive, green As Primitive, blue As Primitive) As Primitive
        Dim num = System.Math.Abs(CInt(red) Mod 256)
        Dim num2 = System.Math.Abs(CInt(green) Mod 256)
        Dim num3 = System.Math.Abs(CInt(blue) Mod 256)
        Return $"#{num:X2}{num2:X2}{num3:X2}"
    End Function

    Friend Function FromString(color As String) As Media.Color
        Try
            Return CType(ColorConverter.ConvertFromString(color), Media.Color)
        Catch
            Return Colors.Black
        End Try
    End Function

    Public Function GetName(color As Primitive) As Primitive
        If ColorNames.ContainsKey(color) Then
            Return ColorNames(color)
        Else
            Return color
        End If

    End Function

    Public ReadOnly Property AliceBlue As Primitive = "#F0F8FF"
    Public ReadOnly Property AntiqueWhite As Primitive = "#FAEBD7"
    Public ReadOnly Property Aqua As Primitive = "#00FFFF"
    Public ReadOnly Property Aquamarine As Primitive = "#7FFFD4"
    Public ReadOnly Property Azure As Primitive = "#F0FFFF"
    Public ReadOnly Property Beige As Primitive = "#F5F5DC"
    Public ReadOnly Property Bisque As Primitive = "#FFE4C4"
    Public ReadOnly Property Black As Primitive = "#000000"
    Public ReadOnly Property BlanchedAlmond As Primitive = "#FFEBCD"
    Public ReadOnly Property Blue As Primitive = "#0000FF"
    Public ReadOnly Property BlueViolet As Primitive = "#8A2BE2"
    Public ReadOnly Property Brown As Primitive = "#A52A2A"
    Public ReadOnly Property BurlyWood As Primitive = "#DEB887"
    Public ReadOnly Property CadetBlue As Primitive = "#5F9EA0"
    Public ReadOnly Property Chartreuse As Primitive = "#7FFF00"
    Public ReadOnly Property Chocolate As Primitive = "#D2691E"
    Public ReadOnly Property Coral As Primitive = "#FF7F50"
    Public ReadOnly Property CornflowerBlue As Primitive = "#6495ED"
    Public ReadOnly Property Cornsilk As Primitive = "#FFF8DC"
    Public ReadOnly Property Crimson As Primitive = "#DC143C"
    Public ReadOnly Property Cyan As Primitive = "#00FFFF"
    Public ReadOnly Property DarkBlue As Primitive = "#00008B"
    Public ReadOnly Property DarkCyan As Primitive = "#008B8B"
    Public ReadOnly Property DarkGoldenrod As Primitive = "#B8860B"
    Public ReadOnly Property DarkGray As Primitive = "#A9A9A9"
    Public ReadOnly Property DarkGreen As Primitive = "#006400"
    Public ReadOnly Property DarkKhaki As Primitive = "#BDB76B"
    Public ReadOnly Property DarkMagenta As Primitive = "#8B008B"
    Public ReadOnly Property DarkOliveGreen As Primitive = "#556B2F"
    Public ReadOnly Property DarkOrange As Primitive = "#FF8C00"
    Public ReadOnly Property DarkOrchid As Primitive = "#9932CC"
    Public ReadOnly Property DarkRed As Primitive = "#8B0000"
    Public ReadOnly Property DarkSalmon As Primitive = "#E9967A"
    Public ReadOnly Property DarkSeaGreen As Primitive = "#8FBC8F"
    Public ReadOnly Property DarkSlateBlue As Primitive = "#483D8B"
    Public ReadOnly Property DarkSlateGray As Primitive = "#2F4F4F"
    Public ReadOnly Property DarkTurquoise As Primitive = "#00CED1"
    Public ReadOnly Property DarkViolet As Primitive = "#9400D3"
    Public ReadOnly Property DeepPink As Primitive = "#FF1493"
    Public ReadOnly Property DeepSkyBlue As Primitive = "#00BFFF"
    Public ReadOnly Property DimGray As Primitive = "#696969"
    Public ReadOnly Property DodgerBlue As Primitive = "#1E90FF"
    Public ReadOnly Property FireBrick As Primitive = "#B22222"
    Public ReadOnly Property FloralWhite As Primitive = "#FFFAF0"
    Public ReadOnly Property ForestGreen As Primitive = "#228B22"
    Public ReadOnly Property Fuchsia As Primitive = "#FF00FF"
    Public ReadOnly Property Gainsboro As Primitive = "#DCDCDC"
    Public ReadOnly Property GhostWhite As Primitive = "#F8F8FF"
    Public ReadOnly Property Gold As Primitive = "#FFD700"
    Public ReadOnly Property Goldenrod As Primitive = "#DAA520"
    Public ReadOnly Property Gray As Primitive = "#808080"
    Public ReadOnly Property Green As Primitive = "#008000"
    Public ReadOnly Property GreenYellow As Primitive = "#ADFF2F"
    Public ReadOnly Property Honeydew As Primitive = "#F0FFF0"
    Public ReadOnly Property HotPink As Primitive = "#FF69B4"
    Public ReadOnly Property IndianRed As Primitive = "#CD5C5C"
    Public ReadOnly Property Indigo As Primitive = "#4B0082"
    Public ReadOnly Property Ivory As Primitive = "#FFFFF0"
    Public ReadOnly Property Khaki As Primitive = "#F0E68C"
    Public ReadOnly Property Lavender As Primitive = "#E6E6FA"
    Public ReadOnly Property LavenderBlush As Primitive = "#FFF0F5"
    Public ReadOnly Property LawnGreen As Primitive = "#7CFC00"
    Public ReadOnly Property LemonChiffon As Primitive = "#FFFACD"
    Public ReadOnly Property LightBlue As Primitive = "#ADD8E6"
    Public ReadOnly Property LightCoral As Primitive = "#F08080"
    Public ReadOnly Property LightCyan As Primitive = "#E0FFFF"
    Public ReadOnly Property LightGoldenrodYellow As Primitive = "#FAFAD2"
    Public ReadOnly Property LightGray As Primitive = "#D3D3D3"
    Public ReadOnly Property LightGreen As Primitive = "#90EE90"
    Public ReadOnly Property LightPink As Primitive = "#FFB6C1"
    Public ReadOnly Property LightSalmon As Primitive = "#FFA07A"
    Public ReadOnly Property LightSeaGreen As Primitive = "#20B2AA"
    Public ReadOnly Property LightSkyBlue As Primitive = "#87CEFA"
    Public ReadOnly Property LightSlateGray As Primitive = "#778899"
    Public ReadOnly Property LightSteelBlue As Primitive = "#B0C4DE"
    Public ReadOnly Property LightYellow As Primitive = "#FFFFE0"
    Public ReadOnly Property Lime As Primitive = "#00FF00"
    Public ReadOnly Property LimeGreen As Primitive = "#32CD32"
    Public ReadOnly Property Linen As Primitive = "#FAF0E6"
    Public ReadOnly Property Magenta As Primitive = "#FF00FF"
    Public ReadOnly Property Maroon As Primitive = "#800000"
    Public ReadOnly Property MediumAquamarine As Primitive = "#66CDAA"
    Public ReadOnly Property MediumBlue As Primitive = "#0000CD"
    Public ReadOnly Property MediumOrchid As Primitive = "#BA55D3"
    Public ReadOnly Property MediumPurple As Primitive = "#9370DB"
    Public ReadOnly Property MediumSeaGreen As Primitive = "#3CB371"
    Public ReadOnly Property MediumSlateBlue As Primitive = "#7B68EE"
    Public ReadOnly Property MediumSpringGreen As Primitive = "#00FA9A"
    Public ReadOnly Property MediumTurquoise As Primitive = "#48D1CC"
    Public ReadOnly Property MediumVioletRed As Primitive = "#C71585"
    Public ReadOnly Property MidnightBlue As Primitive = "#191970"
    Public ReadOnly Property MintCream As Primitive = "#F5FFFA"
    Public ReadOnly Property MistyRose As Primitive = "#FFE4E1"
    Public ReadOnly Property Moccasin As Primitive = "#FFE4B5"
    Public ReadOnly Property NavajoWhite As Primitive = "#FFDEAD"
    Public ReadOnly Property Navy As Primitive = "#000080"
    Public ReadOnly Property OldLace As Primitive = "#FDF5E6"
    Public ReadOnly Property Olive As Primitive = "#808000"
    Public ReadOnly Property OliveDrab As Primitive = "#6B8E23"
    Public ReadOnly Property Orange As Primitive = "#FFA500"
    Public ReadOnly Property OrangeRed As Primitive = "#FF4500"
    Public ReadOnly Property Orchid As Primitive = "#DA70D6"
    Public ReadOnly Property PaleGoldenrod As Primitive = "#EEE8AA"
    Public ReadOnly Property PaleGreen As Primitive = "#98FB98"
    Public ReadOnly Property PaleTurquoise As Primitive = "#AFEEEE"
    Public ReadOnly Property PaleVioletRed As Primitive = "#DB7093"
    Public ReadOnly Property PapayaWhip As Primitive = "#FFEFD5"
    Public ReadOnly Property PeachPuff As Primitive = "#FFDAB9"
    Public ReadOnly Property Peru As Primitive = "#CD853F"
    Public ReadOnly Property Pink As Primitive = "#FFC0CB"
    Public ReadOnly Property Plum As Primitive = "#DDA0DD"
    Public ReadOnly Property PowderBlue As Primitive = "#B0E0E6"
    Public ReadOnly Property Purple As Primitive = "#800080"
    Public ReadOnly Property Red As Primitive = "#FF0000"
    Public ReadOnly Property RosyBrown As Primitive = "#BC8F8F"
    Public ReadOnly Property RoyalBlue As Primitive = "#4169E1"
    Public ReadOnly Property SaddleBrown As Primitive = "#8B4513"
    Public ReadOnly Property Salmon As Primitive = "#FA8072"
    Public ReadOnly Property SandyBrown As Primitive = "#F4A460"
    Public ReadOnly Property SeaGreen As Primitive = "#2E8B57"
    Public ReadOnly Property Seashell As Primitive = "#FFF5EE"
    Public ReadOnly Property Sienna As Primitive = "#A0522D"
    Public ReadOnly Property Silver As Primitive = "#C0C0C0"
    Public ReadOnly Property SkyBlue As Primitive = "#87CEEB"
    Public ReadOnly Property SlateBlue As Primitive = "#6A5ACD"
    Public ReadOnly Property SlateGray As Primitive = "#708090"
    Public ReadOnly Property Snow As Primitive = "#FFFAFA"
    Public ReadOnly Property SpringGreen As Primitive = "#00FF7F"
    Public ReadOnly Property SteelBlue As Primitive = "#4682B4"
    Public ReadOnly Property Tan As Primitive = "#D2B48C"
    Public ReadOnly Property Teal As Primitive = "#008080"
    Public ReadOnly Property Thistle As Primitive = "#D8BFD8"
    Public ReadOnly Property Tomato As Primitive = "#FF6347"
    Public ReadOnly Property Transparent As Primitive = "#00FFFFFF"
    Public ReadOnly Property Turquoise As Primitive = "#40E0D0"
    Public ReadOnly Property Violet As Primitive = "#EE82EE"
    Public ReadOnly Property Wheat As Primitive = "#F5DEB3"
    Public ReadOnly Property White As Primitive = "#FFFFFF"
    Public ReadOnly Property WhiteSmoke As Primitive = "#F5F5F5"
    Public ReadOnly Property Yellow As Primitive = "#FFFF00"
    Public ReadOnly Property YellowGreen As Primitive = "#9ACD32"

    Dim ColorNames As New Dictionary(Of String, String) From {
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


End Module
