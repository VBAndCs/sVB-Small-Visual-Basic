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

    Public Property AliceBlue As Primitive = "#F0F8FF"
    Public Property AntiqueWhite As Primitive = "#FAEBD7"
    Public Property Aqua As Primitive = "#00FFFF"
    Public Property Aquamarine As Primitive = "#7FFFD4"
    Public Property Azure As Primitive = "#F0FFFF"
    Public Property Beige As Primitive = "#F5F5DC"
    Public Property Bisque As Primitive = "#FFE4C4"
    Public Property Black As Primitive = "#000000"
    Public Property BlanchedAlmond As Primitive = "#FFEBCD"
    Public Property Blue As Primitive = "#0000FF"
    Public Property BlueViolet As Primitive = "#8A2BE2"
    Public Property Brown As Primitive = "#A52A2A"
    Public Property BurlyWood As Primitive = "#DEB887"
    Public Property CadetBlue As Primitive = "#5F9EA0"
    Public Property Chartreuse As Primitive = "#7FFF00"
    Public Property Chocolate As Primitive = "#D2691E"
    Public Property Coral As Primitive = "#FF7F50"
    Public Property CornflowerBlue As Primitive = "#6495ED"
    Public Property Cornsilk As Primitive = "#FFF8DC"
    Public Property Crimson As Primitive = "#DC143C"
    Public Property Cyan As Primitive = "#00FFFF"
    Public Property DarkBlue As Primitive = "#00008B"
    Public Property DarkCyan As Primitive = "#008B8B"
    Public Property DarkGoldenrod As Primitive = "#B8860B"
    Public Property DarkGray As Primitive = "#A9A9A9"
    Public Property DarkGreen As Primitive = "#006400"
    Public Property DarkKhaki As Primitive = "#BDB76B"
    Public Property DarkMagenta As Primitive = "#8B008B"
    Public Property DarkOliveGreen As Primitive = "#556B2F"
    Public Property DarkOrange As Primitive = "#FF8C00"
    Public Property DarkOrchid As Primitive = "#9932CC"
    Public Property DarkRed As Primitive = "#8B0000"
    Public Property DarkSalmon As Primitive = "#E9967A"
    Public Property DarkSeaGreen As Primitive = "#8FBC8F"
    Public Property DarkSlateBlue As Primitive = "#483D8B"
    Public Property DarkSlateGray As Primitive = "#2F4F4F"
    Public Property DarkTurquoise As Primitive = "#00CED1"
    Public Property DarkViolet As Primitive = "#9400D3"
    Public Property DeepPink As Primitive = "#FF1493"
    Public Property DeepSkyBlue As Primitive = "#00BFFF"
    Public Property DimGray As Primitive = "#696969"
    Public Property DodgerBlue As Primitive = "#1E90FF"
    Public Property FireBrick As Primitive = "#B22222"
    Public Property FloralWhite As Primitive = "#FFFAF0"
    Public Property ForestGreen As Primitive = "#228B22"
    Public Property Fuchsia As Primitive = "#FF00FF"
    Public Property Gainsboro As Primitive = "#DCDCDC"
    Public Property GhostWhite As Primitive = "#F8F8FF"
    Public Property Gold As Primitive = "#FFD700"
    Public Property Goldenrod As Primitive = "#DAA520"
    Public Property Gray As Primitive = "#808080"
    Public Property Green As Primitive = "#008000"
    Public Property GreenYellow As Primitive = "#ADFF2F"
    Public Property Honeydew As Primitive = "#F0FFF0"
    Public Property HotPink As Primitive = "#FF69B4"
    Public Property IndianRed As Primitive = "#CD5C5C"
    Public Property Indigo As Primitive = "#4B0082"
    Public Property Ivory As Primitive = "#FFFFF0"
    Public Property Khaki As Primitive = "#F0E68C"
    Public Property Lavender As Primitive = "#E6E6FA"
    Public Property LavenderBlush As Primitive = "#FFF0F5"
    Public Property LawnGreen As Primitive = "#7CFC00"
    Public Property LemonChiffon As Primitive = "#FFFACD"
    Public Property LightBlue As Primitive = "#ADD8E6"
    Public Property LightCoral As Primitive = "#F08080"
    Public Property LightCyan As Primitive = "#E0FFFF"
    Public Property LightGoldenrodYellow As Primitive = "#FAFAD2"
    Public Property LightGray As Primitive = "#D3D3D3"
    Public Property LightGreen As Primitive = "#90EE90"
    Public Property LightPink As Primitive = "#FFB6C1"
    Public Property LightSalmon As Primitive = "#FFA07A"
    Public Property LightSeaGreen As Primitive = "#20B2AA"
    Public Property LightSkyBlue As Primitive = "#87CEFA"
    Public Property LightSlateGray As Primitive = "#778899"
    Public Property LightSteelBlue As Primitive = "#B0C4DE"
    Public Property LightYellow As Primitive = "#FFFFE0"
    Public Property Lime As Primitive = "#00FF00"
    Public Property LimeGreen As Primitive = "#32CD32"
    Public Property Linen As Primitive = "#FAF0E6"
    Public Property Magenta As Primitive = "#FF00FF"
    Public Property Maroon As Primitive = "#800000"
    Public Property MediumAquamarine As Primitive = "#66CDAA"
    Public Property MediumBlue As Primitive = "#0000CD"
    Public Property MediumOrchid As Primitive = "#BA55D3"
    Public Property MediumPurple As Primitive = "#9370DB"
    Public Property MediumSeaGreen As Primitive = "#3CB371"
    Public Property MediumSlateBlue As Primitive = "#7B68EE"
    Public Property MediumSpringGreen As Primitive = "#00FA9A"
    Public Property MediumTurquoise As Primitive = "#48D1CC"
    Public Property MediumVioletRed As Primitive = "#C71585"
    Public Property MidnightBlue As Primitive = "#191970"
    Public Property MintCream As Primitive = "#F5FFFA"
    Public Property MistyRose As Primitive = "#FFE4E1"
    Public Property Moccasin As Primitive = "#FFE4B5"
    Public Property NavajoWhite As Primitive = "#FFDEAD"
    Public Property Navy As Primitive = "#000080"
    Public Property OldLace As Primitive = "#FDF5E6"
    Public Property Olive As Primitive = "#808000"
    Public Property OliveDrab As Primitive = "#6B8E23"
    Public Property Orange As Primitive = "#FFA500"
    Public Property OrangeRed As Primitive = "#FF4500"
    Public Property Orchid As Primitive = "#DA70D6"
    Public Property PaleGoldenrod As Primitive = "#EEE8AA"
    Public Property PaleGreen As Primitive = "#98FB98"
    Public Property PaleTurquoise As Primitive = "#AFEEEE"
    Public Property PaleVioletRed As Primitive = "#DB7093"
    Public Property PapayaWhip As Primitive = "#FFEFD5"
    Public Property PeachPuff As Primitive = "#FFDAB9"
    Public Property Peru As Primitive = "#CD853F"
    Public Property Pink As Primitive = "#FFC0CB"
    Public Property Plum As Primitive = "#DDA0DD"
    Public Property PowderBlue As Primitive = "#B0E0E6"
    Public Property Purple As Primitive = "#800080"
    Public Property Red As Primitive = "#FF0000"
    Public Property RosyBrown As Primitive = "#BC8F8F"
    Public Property RoyalBlue As Primitive = "#4169E1"
    Public Property SaddleBrown As Primitive = "#8B4513"
    Public Property Salmon As Primitive = "#FA8072"
    Public Property SandyBrown As Primitive = "#F4A460"
    Public Property SeaGreen As Primitive = "#2E8B57"
    Public Property Seashell As Primitive = "#FFF5EE"
    Public Property Sienna As Primitive = "#A0522D"
    Public Property Silver As Primitive = "#C0C0C0"
    Public Property SkyBlue As Primitive = "#87CEEB"
    Public Property SlateBlue As Primitive = "#6A5ACD"
    Public Property SlateGray As Primitive = "#708090"
    Public Property Snow As Primitive = "#FFFAFA"
    Public Property SpringGreen As Primitive = "#00FF7F"
    Public Property SteelBlue As Primitive = "#4682B4"
    Public Property Tan As Primitive = "#D2B48C"
    Public Property Teal As Primitive = "#008080"
    Public Property Thistle As Primitive = "#D8BFD8"
    Public Property Tomato As Primitive = "#FF6347"
    Public Property Transparent As Primitive = "#00FFFFFF"
    Public Property Turquoise As Primitive = "#40E0D0"
    Public Property Violet As Primitive = "#EE82EE"
    Public Property Wheat As Primitive = "#F5DEB3"
    Public Property White As Primitive = "#FFFFFF"
    Public Property WhiteSmoke As Primitive = "#F5F5F5"
    Public Property Yellow As Primitive = "#FFFF00"
    Public Property YellowGreen As Primitive = "#9ACD32"

    Dim ColorNames As New Dictionary(Of String, String)

    Sub New()
        ColorNames.Add("#F0F8FF", "AliceBlue")
        ColorNames.Add("#FAEBD7", "AntiqueWhite")
        ColorNames.Add("#7FFFD4", "Aquamarine")
        ColorNames.Add("#F0FFFF", "Azure")
        ColorNames.Add("#F5F5DC", "Beige")
        ColorNames.Add("#FFE4C4", "Bisque")
        ColorNames.Add("#000000", "Black")
        ColorNames.Add("#FFEBCD", "BlanchedAlmond")
        ColorNames.Add("#0000FF", "Blue")
        ColorNames.Add("#8A2BE2", "BlueViolet")
        ColorNames.Add("#A52A2A", "Brown")
        ColorNames.Add("#DEB887", "BurlyWood")
        ColorNames.Add("#5F9EA0", "CadetBlue")
        ColorNames.Add("#7FFF00", "Chartreuse")
        ColorNames.Add("#D2691E", "Chocolate")
        ColorNames.Add("#FF7F50", "Coral")
        ColorNames.Add("#6495ED", "CornflowerBlue")
        ColorNames.Add("#FFF8DC", "Cornsilk")
        ColorNames.Add("#DC143C", "Crimson")
        ColorNames.Add("#00FFFF", "Cyan")
        ColorNames.Add("#00008B", "DarkBlue")
        ColorNames.Add("#008B8B", "DarkCyan")
        ColorNames.Add("#B8860B", "DarkGoldenrod")
        ColorNames.Add("#A9A9A9", "DarkGray")
        ColorNames.Add("#006400", "DarkGreen")
        ColorNames.Add("#BDB76B", "DarkKhaki")
        ColorNames.Add("#8B008B", "DarkMagenta")
        ColorNames.Add("#556B2F", "DarkOliveGreen")
        ColorNames.Add("#FF8C00", "DarkOrange")
        ColorNames.Add("#9932CC", "DarkOrchid")
        ColorNames.Add("#8B0000", "DarkRed")
        ColorNames.Add("#E9967A", "DarkSalmon")
        ColorNames.Add("#8FBC8F", "DarkSeaGreen")
        ColorNames.Add("#483D8B", "DarkSlateBlue")
        ColorNames.Add("#2F4F4F", "DarkSlateGray")
        ColorNames.Add("#00CED1", "DarkTurquoise")
        ColorNames.Add("#9400D3", "DarkViolet")
        ColorNames.Add("#FF1493", "DeepPink")
        ColorNames.Add("#00BFFF", "DeepSkyBlue")
        ColorNames.Add("#696969", "DimGray")
        ColorNames.Add("#1E90FF", "DodgerBlue")
        ColorNames.Add("#B22222", "FireBrick")
        ColorNames.Add("#FFFAF0", "FloralWhite")
        ColorNames.Add("#228B22", "ForestGreen")
        ColorNames.Add("#FF00FF", "Fuchsia")
        ColorNames.Add("#DCDCDC", "Gainsboro")
        ColorNames.Add("#F8F8FF", "GhostWhite")
        ColorNames.Add("#FFD700", "Gold")
        ColorNames.Add("#DAA520", "Goldenrod")
        ColorNames.Add("#808080", "Gray")
        ColorNames.Add("#008000", "Green")
        ColorNames.Add("#ADFF2F", "GreenYellow")
        ColorNames.Add("#F0FFF0", "Honeydew")
        ColorNames.Add("#FF69B4", "HotPink")
        ColorNames.Add("#CD5C5C", "IndianRed")
        ColorNames.Add("#4B0082", "Indigo")
        ColorNames.Add("#FFFFF0", "Ivory")
        ColorNames.Add("#F0E68C", "Khaki")
        ColorNames.Add("#E6E6FA", "Lavender")
        ColorNames.Add("#FFF0F5", "LavenderBlush")
        ColorNames.Add("#7CFC00", "LawnGreen")
        ColorNames.Add("#FFFACD", "LemonChiffon")
        ColorNames.Add("#ADD8E6", "LightBlue")
        ColorNames.Add("#F08080", "LightCoral")
        ColorNames.Add("#E0FFFF", "LightCyan")
        ColorNames.Add("#FAFAD2", "LightGoldenrodYellow")
        ColorNames.Add("#D3D3D3", "LightGray")
        ColorNames.Add("#90EE90", "LightGreen")
        ColorNames.Add("#FFB6C1", "LightPink")
        ColorNames.Add("#FFA07A", "LightSalmon")
        ColorNames.Add("#20B2AA", "LightSeaGreen")
        ColorNames.Add("#87CEFA", "LightSkyBlue")
        ColorNames.Add("#778899", "LightSlateGray")
        ColorNames.Add("#B0C4DE", "LightSteelBlue")
        ColorNames.Add("#FFFFE0", "LightYellow")
        ColorNames.Add("#00FF00", "Lime")
        ColorNames.Add("#32CD32", "LimeGreen")
        ColorNames.Add("#FAF0E6", "Linen")
        ColorNames.Add("#800000", "Maroon")
        ColorNames.Add("#66CDAA", "MediumAquamarine")
        ColorNames.Add("#0000CD", "MediumBlue")
        ColorNames.Add("#BA55D3", "MediumOrchid")
        ColorNames.Add("#9370DB", "MediumPurple")
        ColorNames.Add("#3CB371", "MediumSeaGreen")
        ColorNames.Add("#7B68EE", "MediumSlateBlue")
        ColorNames.Add("#00FA9A", "MediumSpringGreen")
        ColorNames.Add("#48D1CC", "MediumTurquoise")
        ColorNames.Add("#C71585", "MediumVioletRed")
        ColorNames.Add("#191970", "MidnightBlue")
        ColorNames.Add("#F5FFFA", "MintCream")
        ColorNames.Add("#FFE4E1", "MistyRose")
        ColorNames.Add("#FFE4B5", "Moccasin")
        ColorNames.Add("#FFDEAD", "NavajoWhite")
        ColorNames.Add("#000080", "Navy")
        ColorNames.Add("#FDF5E6", "OldLace")
        ColorNames.Add("#808000", "Olive")
        ColorNames.Add("#6B8E23", "OliveDrab")
        ColorNames.Add("#FFA500", "Orange")
        ColorNames.Add("#FF4500", "OrangeRed")
        ColorNames.Add("#DA70D6", "Orchid")
        ColorNames.Add("#EEE8AA", "PaleGoldenrod")
        ColorNames.Add("#98FB98", "PaleGreen")
        ColorNames.Add("#AFEEEE", "PaleTurquoise")
        ColorNames.Add("#DB7093", "PaleVioletRed")
        ColorNames.Add("#FFEFD5", "PapayaWhip")
        ColorNames.Add("#FFDAB9", "PeachPuff")
        ColorNames.Add("#CD853F", "Peru")
        ColorNames.Add("#FFC0CB", "Pink")
        ColorNames.Add("#DDA0DD", "Plum")
        ColorNames.Add("#B0E0E6", "PowderBlue")
        ColorNames.Add("#800080", "Purple")
        ColorNames.Add("#FF0000", "Red")
        ColorNames.Add("#BC8F8F", "RosyBrown")
        ColorNames.Add("#4169E1", "RoyalBlue")
        ColorNames.Add("#8B4513", "SaddleBrown")
        ColorNames.Add("#FA8072", "Salmon")
        ColorNames.Add("#F4A460", "SandyBrown")
        ColorNames.Add("#2E8B57", "SeaGreen")
        ColorNames.Add("#FFF5EE", "Seashell")
        ColorNames.Add("#A0522D", "Sienna")
        ColorNames.Add("#C0C0C0", "Silver")
        ColorNames.Add("#87CEEB", "SkyBlue")
        ColorNames.Add("#6A5ACD", "SlateBlue")
        ColorNames.Add("#708090", "SlateGray")
        ColorNames.Add("#FFFAFA", "Snow")
        ColorNames.Add("#00FF7F", "SpringGreen")
        ColorNames.Add("#4682B4", "SteelBlue")
        ColorNames.Add("#D2B48C", "Tan")
        ColorNames.Add("#008080", "Teal")
        ColorNames.Add("#D8BFD8", "Thistle")
        ColorNames.Add("#FF6347", "Tomato")
        ColorNames.Add("#00FFFFFF", "Transparent")
        ColorNames.Add("#40E0D0", "Turquoise")
        ColorNames.Add("#EE82EE", "Violet")
        ColorNames.Add("#F5DEB3", "Wheat")
        ColorNames.Add("#FFFFFF", "White")
        ColorNames.Add("#F5F5F5", "WhiteSmoke")
        ColorNames.Add("#FFFF00", "Yellow")
        ColorNames.Add("#9ACD32", "YellowGreen")
    End Sub

End Module
