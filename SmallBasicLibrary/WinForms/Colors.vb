Imports Microsoft.SmallVisualBasic.Library
Imports SysColors = System.Windows.SystemColors

Namespace WinForms
    ''' <summary>Defines all known color names</summary>
    <SmallVisualBasicType>
    Public NotInheritable Class Colors
        ''' <summary>
        ''' AliceBlue Color:
        ''' Hex: "#FFF0F8FF"
        ''' R=240, G=248, B=255
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property AliceBlue As Primitive = "#FFF0F8FF"

        ''' <summary>
        ''' AntiqueWhite Color:
        ''' Hex: "#FFFAEBD7"
        ''' R=250, G=235, B=215
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property AntiqueWhite As Primitive = "#FFFAEBD7"

        ''' <summary>
        ''' Cyan Color:
        ''' Hex: "#FF00FFFF"
        ''' R=0, G=255, B=255
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property Aqua As Primitive = "#FF00FFFF"

        ''' <summary>
        ''' Aquamarine Color:
        ''' Hex: "#FF7FFFD4"
        ''' R=127, G=255, B=212
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property Aquamarine As Primitive = "#FF7FFFD4"

        ''' <summary>
        ''' Azure Color:
        ''' Hex: "#FFF0FFFF"
        ''' R=240, G=255, B=255
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property Azure As Primitive = "#FFF0FFFF"

        ''' <summary>
        ''' Beige Color:
        ''' Hex: "#FFF5F5DC"
        ''' R=245, G=245, B=220
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property Beige As Primitive = "#FFF5F5DC"

        ''' <summary>
        ''' Bisque Color:
        ''' Hex: "#FFFFE4C4"
        ''' R=255, G=228, B=196
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property Bisque As Primitive = "#FFFFE4C4"

        ''' <summary>
        ''' Black Color:
        ''' Hex: "#FF000000"
        ''' R=0, G=0, B=0
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property Black As Primitive = "#FF000000"

        ''' <summary>
        ''' BlanchedAlmond Color:
        ''' Hex: "#FFFFEBCD"
        ''' R=255, G=235, B=205
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property BlanchedAlmond As Primitive = "#FFFFEBCD"

        ''' <summary>
        ''' Blue Color:
        ''' Hex: "#FF0000FF"
        ''' R=0, G=0, B=255
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property Blue As Primitive = "#FF0000FF"

        ''' <summary>
        ''' BlueViolet Color:
        ''' Hex: "#FF8A2BE2"
        ''' R=138, G=43, B=226
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property BlueViolet As Primitive = "#FF8A2BE2"

        ''' <summary>
        ''' Brown Color:
        ''' Hex: "#FFA52A2A"
        ''' R=165, G=42, B=42
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property Brown As Primitive = "#FFA52A2A"

        ''' <summary>
        ''' BurlyWood Color:
        ''' Hex: "#FFDEB887"
        ''' R=222, G=184, B=135
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property BurlyWood As Primitive = "#FFDEB887"

        ''' <summary>
        ''' CadetBlue Color:
        ''' Hex: "#FF5F9EA0"
        ''' R=95, G=158, B=160
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property CadetBlue As Primitive = "#FF5F9EA0"

        ''' <summary>
        ''' Chartreuse Color:
        ''' Hex: "#FF7FFF00"
        ''' R=127, G=255, B=0
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property Chartreuse As Primitive = "#FF7FFF00"

        ''' <summary>
        ''' Chocolate Color:
        ''' Hex: "#FFD2691E"
        ''' R=210, G=105, B=30
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property Chocolate As Primitive = "#FFD2691E"

        ''' <summary>
        ''' Coral Color:
        ''' Hex: "#FFFF7F50"
        ''' R=255, G=127, B=80
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property Coral As Primitive = "#FFFF7F50"

        ''' <summary>
        ''' CornflowerBlue Color:
        ''' Hex: "#FF6495ED"
        ''' R=100, G=149, B=237
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property CornflowerBlue As Primitive = "#FF6495ED"

        ''' <summary>
        ''' Cornsilk Color:
        ''' Hex: "#FFFFF8DC"
        ''' R=255, G=248, B=220
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property Cornsilk As Primitive = "#FFFFF8DC"

        ''' <summary>
        ''' Crimson Color:
        ''' Hex: "#FFDC143C"
        ''' R=220, G=20, B=60
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property Crimson As Primitive = "#FFDC143C"

        ''' <summary>
        ''' Cyan Color:
        ''' Hex: "#FF00FFFF"
        ''' R=0, G=255, B=255
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property Cyan As Primitive = "#FF00FFFF"

        ''' <summary>
        ''' DarkBlue Color:
        ''' Hex: "#FF00008B"
        ''' R=0, G=0, B=139
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property DarkBlue As Primitive = "#FF00008B"

        ''' <summary>
        ''' DarkCyan Color:
        ''' Hex: "#FF008B8B"
        ''' R=0, G=139, B=139
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property DarkCyan As Primitive = "#FF008B8B"

        ''' <summary>
        ''' DarkGoldenrod Color:
        ''' Hex: "#FFB8860B"
        ''' R=184, G=134, B=11
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property DarkGoldenrod As Primitive = "#FFB8860B"

        ''' <summary>
        ''' DarkGray Color:
        ''' Hex: "#FFA9A9A9"
        ''' R=169, G=169, B=169
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property DarkGray As Primitive = "#FFA9A9A9"

        ''' <summary>
        ''' DarkGreen Color:
        ''' Hex: "#FF006400"
        ''' R=0, G=100, B=0
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property DarkGreen As Primitive = "#FF006400"

        ''' <summary>
        ''' DarkKhaki Color:
        ''' Hex: "#FFBDB76B"
        ''' R=189, G=183, B=107
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property DarkKhaki As Primitive = "#FFBDB76B"

        ''' <summary>
        ''' DarkMagenta Color:
        ''' Hex: "#FF8B008B"
        ''' R=139, G=0, B=139
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property DarkMagenta As Primitive = "#FF8B008B"

        ''' <summary>
        ''' DarkOliveGreen Color:
        ''' Hex: "#FF556B2F"
        ''' R=85, G=107, B=47
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property DarkOliveGreen As Primitive = "#FF556B2F"

        ''' <summary>
        ''' DarkOrange Color:
        ''' Hex: "#FFFF8C00"
        ''' R=255, G=140, B=0
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property DarkOrange As Primitive = "#FFFF8C00"

        ''' <summary>
        ''' DarkOrchid Color:
        ''' Hex: "#FF9932CC"
        ''' R=153, G=50, B=204
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property DarkOrchid As Primitive = "#FF9932CC"

        ''' <summary>
        ''' DarkRed Color:
        ''' Hex: "#FF8B0000"
        ''' R=139, G=0, B=0
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property DarkRed As Primitive = "#FF8B0000"

        ''' <summary>
        ''' DarkSalmon Color:
        ''' Hex: "#FFE9967A"
        ''' R=233, G=150, B=122
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property DarkSalmon As Primitive = "#FFE9967A"

        ''' <summary>
        ''' DarkSeaGreen Color:
        ''' Hex: "#FF8FBC8F"
        ''' R=143, G=188, B=143
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property DarkSeaGreen As Primitive = "#FF8FBC8F"

        ''' <summary>
        ''' DarkSlateBlue Color:
        ''' Hex: "#FF483D8B"
        ''' R=72, G=61, B=139
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property DarkSlateBlue As Primitive = "#FF483D8B"

        ''' <summary>
        ''' DarkSlateGray Color:
        ''' Hex: "#FF2F4F4F"
        ''' R=47, G=79, B=79
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property DarkSlateGray As Primitive = "#FF2F4F4F"

        ''' <summary>
        ''' DarkTurquoise Color:
        ''' Hex: "#FF00CED1"
        ''' R=0, G=206, B=209
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property DarkTurquoise As Primitive = "#FF00CED1"

        ''' <summary>
        ''' DarkViolet Color:
        ''' Hex: "#FF9400D3"
        ''' R=148, G=0, B=211
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property DarkViolet As Primitive = "#FF9400D3"

        ''' <summary>
        ''' DeepPink Color:
        ''' Hex: "#FFFF1493"
        ''' R=255, G=20, B=147
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property DeepPink As Primitive = "#FFFF1493"

        ''' <summary>
        ''' DeepSkyBlue Color:
        ''' Hex: "#FF00BFFF"
        ''' R=0, G=191, B=255
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property DeepSkyBlue As Primitive = "#FF00BFFF"

        ''' <summary>
        ''' DimGray Color:
        ''' Hex: "#FF696969"
        ''' R=105, G=105, B=105
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property DimGray As Primitive = "#FF696969"

        ''' <summary>
        ''' DodgerBlue Color:
        ''' Hex: "#FF1E90FF"
        ''' R=30, G=144, B=255
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property DodgerBlue As Primitive = "#FF1E90FF"

        ''' <summary>
        ''' FireBrick Color:
        ''' Hex: "#FFB22222"
        ''' R=178, G=34, B=34
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property FireBrick As Primitive = "#FFB22222"

        ''' <summary>
        ''' FloralWhite Color:
        ''' Hex: "#FFFFFAF0"
        ''' R=255, G=250, B=240
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property FloralWhite As Primitive = "#FFFFFAF0"

        ''' <summary>
        ''' ForestGreen Color:
        ''' Hex: "#FF228B22"
        ''' R=34, G=139, B=34
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property ForestGreen As Primitive = "#FF228B22"

        ''' <summary>
        ''' Fuchsia Color:
        ''' Hex: "#FFFF00FF"
        ''' R=255, G=0, B=255
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property Magenta As Primitive = "#FFFF00FF"

        ''' <summary>
        ''' Gainsboro Color:
        ''' Hex: "#FFDCDCDC"
        ''' R=220, G=220, B=220
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property Gainsboro As Primitive = "#FFDCDCDC"

        ''' <summary>
        ''' GhostWhite Color:
        ''' Hex: "#FFF8F8FF"
        ''' R=248, G=248, B=255
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property GhostWhite As Primitive = "#FFF8F8FF"

        ''' <summary>
        ''' Gold Color:
        ''' Hex: "#FFFFD700"
        ''' R=255, G=215, B=0
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property Gold As Primitive = "#FFFFD700"

        ''' <summary>
        ''' Goldenrod Color:
        ''' Hex: "#FFDAA520"
        ''' R=218, G=165, B=32
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property Goldenrod As Primitive = "#FFDAA520"

        ''' <summary>
        ''' Gray Color:
        ''' Hex: "#FF808080"
        ''' R=128, G=128, B=128
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property Gray As Primitive = "#FF808080"

        ''' <summary>
        ''' Green Color:
        ''' Hex: "#FF008000"
        ''' R=0, G=128, B=0
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property Green As Primitive = "#FF008000"

        ''' <summary>
        ''' GreenYellow Color:
        ''' Hex: "#FFADFF2F"
        ''' R=173, G=255, B=47
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property GreenYellow As Primitive = "#FFADFF2F"

        ''' <summary>
        ''' Honeydew Color:
        ''' Hex: "#FFF0FFF0"
        ''' R=240, G=255, B=240
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property Honeydew As Primitive = "#FFF0FFF0"

        ''' <summary>
        ''' HotPink Color:
        ''' Hex: "#FFFF69B4"
        ''' R=255, G=105, B=180
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property HotPink As Primitive = "#FFFF69B4"

        ''' <summary>
        ''' IndianRed Color:
        ''' Hex: "#FFCD5C5C"
        ''' R=205, G=92, B=92
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property IndianRed As Primitive = "#FFCD5C5C"

        ''' <summary>
        ''' Indigo Color:
        ''' Hex: "#FF4B0082"
        ''' R=75, G=0, B=130
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property Indigo As Primitive = "#FF4B0082"

        ''' <summary>
        ''' Ivory Color:
        ''' Hex: "#FFFFFFF0"
        ''' R=255, G=255, B=240
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property Ivory As Primitive = "#FFFFFFF0"

        ''' <summary>
        ''' Khaki Color:
        ''' Hex: "#FFF0E68C"
        ''' R=240, G=230, B=140
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property Khaki As Primitive = "#FFF0E68C"

        ''' <summary>
        ''' Lavender Color:
        ''' Hex: "#FFE6E6FA"
        ''' R=230, G=230, B=250
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property Lavender As Primitive = "#FFE6E6FA"

        ''' <summary>
        ''' LavenderBlush Color:
        ''' Hex: "#FFFFF0F5"
        ''' R=255, G=240, B=245
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property LavenderBlush As Primitive = "#FFFFF0F5"

        ''' <summary>
        ''' LawnGreen Color:
        ''' Hex: "#FF7CFC00"
        ''' R=124, G=252, B=0
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property LawnGreen As Primitive = "#FF7CFC00"

        ''' <summary>
        ''' LemonChiffon Color:
        ''' Hex: "#FFFFFACD"
        ''' R=255, G=250, B=205
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property LemonChiffon As Primitive = "#FFFFFACD"

        ''' <summary>
        ''' LightBlue Color:
        ''' Hex: "#FFADD8E6"
        ''' R=173, G=216, B=230
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property LightBlue As Primitive = "#FFADD8E6"

        ''' <summary>
        ''' LightCoral Color:
        ''' Hex: "#FFF08080"
        ''' R=240, G=128, B=128
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property LightCoral As Primitive = "#FFF08080"

        ''' <summary>
        ''' LightCyan Color:
        ''' Hex: "#FFE0FFFF"
        ''' R=224, G=255, B=255
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property LightCyan As Primitive = "#FFE0FFFF"

        ''' <summary>
        ''' LightGoldenrodYellow Color:
        ''' Hex: "#FFFAFAD2"
        ''' R=250, G=250, B=210
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property LightGoldenrodYellow As Primitive = "#FFFAFAD2"

        ''' <summary>
        ''' LightGray Color:
        ''' Hex: "#FFD3D3D3"
        ''' R=211, G=211, B=211
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property LightGray As Primitive = "#FFD3D3D3"

        ''' <summary>
        ''' LightGreen Color:
        ''' Hex: "#FF90EE90"
        ''' R=144, G=238, B=144
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property LightGreen As Primitive = "#FF90EE90"

        ''' <summary>
        ''' LightPink Color:
        ''' Hex: "#FFFFB6C1"
        ''' R=255, G=182, B=193
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property LightPink As Primitive = "#FFFFB6C1"

        ''' <summary>
        ''' LightSalmon Color:
        ''' Hex: "#FFFFA07A"
        ''' R=255, G=160, B=122
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property LightSalmon As Primitive = "#FFFFA07A"

        ''' <summary>
        ''' LightSeaGreen Color:
        ''' Hex: "#FF20B2AA"
        ''' R=32, G=178, B=170
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property LightSeaGreen As Primitive = "#FF20B2AA"

        ''' <summary>
        ''' LightSkyBlue Color:
        ''' Hex: "#FF87CEFA"
        ''' R=135, G=206, B=250
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property LightSkyBlue As Primitive = "#FF87CEFA"

        ''' <summary>
        ''' LightSlateGray Color:
        ''' Hex: "#FF778899"
        ''' R=119, G=136, B=153
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property LightSlateGray As Primitive = "#FF778899"

        ''' <summary>
        ''' LightSteelBlue Color:
        ''' Hex: "#FFB0C4DE"
        ''' R=176, G=196, B=222
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property LightSteelBlue As Primitive = "#FFB0C4DE"

        ''' <summary>
        ''' LightYellow Color:
        ''' Hex: "#FFFFFFE0"
        ''' R=255, G=255, B=224
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property LightYellow As Primitive = "#FFFFFFE0"

        ''' <summary>
        ''' Lime Color:
        ''' Hex: "#FF00FF00"
        ''' R=0, G=255, B=0
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property Lime As Primitive = "#FF00FF00"

        ''' <summary>
        ''' LimeGreen Color:
        ''' Hex: "#FF32CD32"
        ''' R=50, G=205, B=50
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property LimeGreen As Primitive = "#FF32CD32"

        ''' <summary>
        ''' Linen Color:
        ''' Hex: "#FFFAF0E6"
        ''' R=250, G=240, B=230
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property Linen As Primitive = "#FFFAF0E6"

        ''' <summary>
        ''' Fuchsia Color:
        ''' Hex: "#FFFF00FF"
        ''' R=255, G=0, B=255
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property Fuchsia As Primitive = "#FFFF00FF"

        ''' <summary>
        ''' Maroon Color:
        ''' Hex: "#FF800000"
        ''' R=128, G=0, B=0
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property Maroon As Primitive = "#FF800000"

        ''' <summary>
        ''' MediumAquamarine Color:
        ''' Hex: "#FF66CDAA"
        ''' R=102, G=205, B=170
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property MediumAquamarine As Primitive = "#FF66CDAA"

        ''' <summary>
        ''' MediumBlue Color:
        ''' Hex: "#FF0000CD"
        ''' R=0, G=0, B=205
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property MediumBlue As Primitive = "#FF0000CD"

        ''' <summary>
        ''' MediumOrchid Color:
        ''' Hex: "#FFBA55D3"
        ''' R=186, G=85, B=211
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property MediumOrchid As Primitive = "#FFBA55D3"

        ''' <summary>
        ''' MediumPurple Color:
        ''' Hex: "#FF9370DB"
        ''' R=147, G=112, B=219
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property MediumPurple As Primitive = "#FF9370DB"

        ''' <summary>
        ''' MediumSeaGreen Color:
        ''' Hex: "#FF3CB371"
        ''' R=60, G=179, B=113
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property MediumSeaGreen As Primitive = "#FF3CB371"

        ''' <summary>
        ''' MediumSlateBlue Color:
        ''' Hex: "#FF7B68EE"
        ''' R=123, G=104, B=238
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property MediumSlateBlue As Primitive = "#FF7B68EE"

        ''' <summary>
        ''' MediumSpringGreen Color:
        ''' Hex: "#FF00FA9A"
        ''' R=0, G=250, B=154
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property MediumSpringGreen As Primitive = "#FF00FA9A"

        ''' <summary>
        ''' MediumTurquoise Color:
        ''' Hex: "#FF48D1CC"
        ''' R=72, G=209, B=204
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property MediumTurquoise As Primitive = "#FF48D1CC"

        ''' <summary>
        ''' MediumVioletRed Color:
        ''' Hex: "#FFC71585"
        ''' R=199, G=21, B=133
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property MediumVioletRed As Primitive = "#FFC71585"

        ''' <summary>
        ''' MidnightBlue Color:
        ''' Hex: "#FF191970"
        ''' R=25, G=25, B=112
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property MidnightBlue As Primitive = "#FF191970"

        ''' <summary>
        ''' MintCream Color:
        ''' Hex: "#FFF5FFFA"
        ''' R=245, G=255, B=250
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property MintCream As Primitive = "#FFF5FFFA"

        ''' <summary>
        ''' MistyRose Color:
        ''' Hex: "#FFFFE4E1"
        ''' R=255, G=228, B=225
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property MistyRose As Primitive = "#FFFFE4E1"

        ''' <summary>
        ''' Moccasin Color:
        ''' Hex: "#FFFFE4B5"
        ''' R=255, G=228, B=181
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property Moccasin As Primitive = "#FFFFE4B5"

        ''' <summary>
        ''' NavajoWhite Color:
        ''' Hex: "#FFFFDEAD"
        ''' R=255, G=222, B=173
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property NavajoWhite As Primitive = "#FFFFDEAD"

        ''' <summary>
        ''' Navy Color:
        ''' Hex: "#FF000080"
        ''' R=0, G=0, B=128
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property Navy As Primitive = "#FF000080"

        ''' <summary>
        ''' No color. Use this value when you don't want to draw the background color or the outline color.
        ''' There is a difference between Colors.None and Colors.Transparent:
        ''' ● The None color deletes the surface of the graphic, so, it doesn't respond to mouse and keyboard events, which are delivered to the underneath control.
        ''' ● The Transparent color keeps the surfuce of the graphic but you can see through it, while it Is still responding to  mouse And keyboard events.
        ''' </summary>
        Public Shared ReadOnly Property None As Primitive = "None"

        ''' <summary>
        ''' OldLace Color:
        ''' Hex: "#FFFDF5E6"
        ''' R=253, G=245, B=230
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property OldLace As Primitive = "#FFFDF5E6"

        ''' <summary>
        ''' Olive Color:
        ''' Hex: "#FF808000"
        ''' R=128, G=128, B=0
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property Olive As Primitive = "#FF808000"

        ''' <summary>
        ''' OliveDrab Color:
        ''' Hex: "#FF6B8E23"
        ''' R=107, G=142, B=35
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property OliveDrab As Primitive = "#FF6B8E23"

        ''' <summary>
        ''' Orange Color:
        ''' Hex: "#FFFFA500"
        ''' R=255, G=165, B=0
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property Orange As Primitive = "#FFFFA500"

        ''' <summary>
        ''' OrangeRed Color:
        ''' Hex: "#FFFF4500"
        ''' R=255, G=69, B=0
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property OrangeRed As Primitive = "#FFFF4500"

        ''' <summary>
        ''' Orchid Color:
        ''' Hex: "#FFDA70D6"
        ''' R=218, G=112, B=214
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property Orchid As Primitive = "#FFDA70D6"

        ''' <summary>
        ''' PaleGoldenrod Color:
        ''' Hex: "#FFEEE8AA"
        ''' R=238, G=232, B=170
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property PaleGoldenrod As Primitive = "#FFEEE8AA"

        ''' <summary>
        ''' PaleGreen Color:
        ''' Hex: "#FF98FB98"
        ''' R=152, G=251, B=152
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property PaleGreen As Primitive = "#FF98FB98"

        ''' <summary>
        ''' PaleTurquoise Color:
        ''' Hex: "#FFAFEEEE"
        ''' R=175, G=238, B=238
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property PaleTurquoise As Primitive = "#FFAFEEEE"

        ''' <summary>
        ''' PaleVioletRed Color:
        ''' Hex: "#FFDB7093"
        ''' R=219, G=112, B=147
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property PaleVioletRed As Primitive = "#FFDB7093"

        ''' <summary>
        ''' PapayaWhip Color:
        ''' Hex: "#FFFFEFD5"
        ''' R=255, G=239, B=213
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property PapayaWhip As Primitive = "#FFFFEFD5"

        ''' <summary>
        ''' PeachPuff Color:
        ''' Hex: "#FFFFDAB9"
        ''' R=255, G=218, B=185
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property PeachPuff As Primitive = "#FFFFDAB9"

        ''' <summary>
        ''' Peru Color:
        ''' Hex: "#FFCD853F"
        ''' R=205, G=133, B=63
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property Peru As Primitive = "#FFCD853F"

        ''' <summary>
        ''' Pink Color:
        ''' Hex: "#FFFFC0CB"
        ''' R=255, G=192, B=203
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property Pink As Primitive = "#FFFFC0CB"

        ''' <summary>
        ''' Plum Color:
        ''' Hex: "#FFDDA0DD"
        ''' R=221, G=160, B=221
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property Plum As Primitive = "#FFDDA0DD"

        ''' <summary>
        ''' PowderBlue Color:
        ''' Hex: "#FFB0E0E6"
        ''' R=176, G=224, B=230
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property PowderBlue As Primitive = "#FFB0E0E6"

        ''' <summary>
        ''' Purple Color:
        ''' Hex: "#FF800080"
        ''' R=128, G=0, B=128
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property Purple As Primitive = "#FF800080"

        ''' <summary>
        ''' Red Color:
        ''' Hex: "#FFFF0000"
        ''' R=255, G=0, B=0
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property Red As Primitive = "#FFFF0000"

        ''' <summary>
        ''' RosyBrown Color:
        ''' Hex: "#FFBC8F8F"
        ''' R=188, G=143, B=143
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property RosyBrown As Primitive = "#FFBC8F8F"

        ''' <summary>
        ''' RoyalBlue Color:
        ''' Hex: "#FF4169E1"
        ''' R=65, G=105, B=225
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property RoyalBlue As Primitive = "#FF4169E1"

        ''' <summary>
        ''' SaddleBrown Color:
        ''' Hex: "#FF8B4513"
        ''' R=139, G=69, B=19
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property SaddleBrown As Primitive = "#FF8B4513"

        ''' <summary>
        ''' Salmon Color:
        ''' Hex: "#FFFA8072"
        ''' R=250, G=128, B=114
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property Salmon As Primitive = "#FFFA8072"

        ''' <summary>
        ''' SandyBrown Color:
        ''' Hex: "#FFF4A460"
        ''' R=244, G=164, B=96
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property SandyBrown As Primitive = "#FFF4A460"

        ''' <summary>
        ''' SeaGreen Color:
        ''' Hex: "#FF2E8B57"
        ''' R=46, G=139, B=87
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property SeaGreen As Primitive = "#FF2E8B57"

        ''' <summary>
        ''' Seashell Color:
        ''' Hex: "#FFFFF5EE"
        ''' R=255, G=245, B=238
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property Seashell As Primitive = "#FFFFF5EE"

        ''' <summary>
        ''' Sienna Color:
        ''' Hex: "#FFA0522D"
        ''' R=160, G=82, B=45
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property Sienna As Primitive = "#FFA0522D"

        ''' <summary>
        ''' Silver Color:
        ''' Hex: "#FFC0C0C0"
        ''' R=192, G=192, B=192
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property Silver As Primitive = "#FFC0C0C0"

        ''' <summary>
        ''' SkyBlue Color:
        ''' Hex: "#FF87CEEB"
        ''' R=135, G=206, B=235
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property SkyBlue As Primitive = "#FF87CEEB"

        ''' <summary>
        ''' SlateBlue Color:
        ''' Hex: "#FF6A5ACD"
        ''' R=106, G=90, B=205
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property SlateBlue As Primitive = "#FF6A5ACD"

        ''' <summary>
        ''' SlateGray Color:
        ''' Hex: "#FF708090"
        ''' R=112, G=128, B=144
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property SlateGray As Primitive = "#FF708090"

        ''' <summary>
        ''' Snow Color:
        ''' Hex: "#FFFFFAFA"
        ''' R=255, G=250, B=250
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property Snow As Primitive = "#FFFFFAFA"

        ''' <summary>
        ''' SpringGreen Color:
        ''' Hex: "#FF00FF7F"
        ''' R=0, G=255, B=127
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property SpringGreen As Primitive = "#FF00FF7F"

        ''' <summary>
        ''' SteelBlue Color:
        ''' Hex: "#FF4682B4"
        ''' R=70, G=130, B=180
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property SteelBlue As Primitive = "#FF4682B4"

        ''' <summary>
        ''' Tan Color:
        ''' Hex: "#FFD2B48C"
        ''' R=210, G=180, B=140
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property Tan As Primitive = "#FFD2B48C"

        ''' <summary>
        ''' Teal Color:
        ''' Hex: "#FF008080"
        ''' R=0, G=128, B=128
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property Teal As Primitive = "#FF008080"

        ''' <summary>
        ''' Thistle Color:
        ''' Hex: "#FFD8BFD8"
        ''' R=216, G=191, B=216
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property Thistle As Primitive = "#FFD8BFD8"

        ''' <summary>
        ''' Tomato Color:
        ''' Hex: "#FFFF6347"
        ''' R=255, G=99, B=71
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property Tomato As Primitive = "#FFFF6347"

        ''' <summary>
        ''' Transparent Color:
        ''' Hex: "#00FFFFFF"
        ''' Alpha = 0, R=255, G=255, B=255
        ''' There is a difference between Colors.Transparent and Colors.None:
        ''' ● The None color deletes the surface of the graphic, so, it doesn't respond to mouse and keyboard events, which are delivered to the underneath control.
        ''' ● The Transparent color keeps the surfuce of the graphic but you can see through it, while it Is still responding to  mouse And keyboard events.
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property Transparent As Primitive = "#FF00FFFFFF"

        ''' <summary>
        ''' Turquoise Color:
        ''' Hex: "#FF40E0D0"
        ''' R=64, G=224, B=208
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property Turquoise As Primitive = "#FF40E0D0"

        ''' <summary>
        ''' Violet Color:
        ''' Hex: "#FFEE82EE"
        ''' R=238, G=130, B=238
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property Violet As Primitive = "#FFEE82EE"

        ''' <summary>
        ''' Wheat Color:
        ''' Hex: "#FFF5DEB3"
        ''' R=245, G=222, B=179
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property Wheat As Primitive = "#FFF5DEB3"

        ''' <summary>
        ''' White Color:
        ''' Hex: "#FFFFFFFF"
        ''' R=255, G=255, B=255
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property White As Primitive = "#FFFFFFFF"

        ''' <summary>
        ''' WhiteSmoke Color:
        ''' Hex: "#FFF5F5F5"
        ''' R=245, G=245, B=245
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property WhiteSmoke As Primitive = "#FFF5F5F5"

        ''' <summary>
        ''' Yellow Color:
        ''' Hex: "#FFFFFF00"
        ''' R=255, G=255, B=0
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property Yellow As Primitive = "#FFFFFF00"

        ''' <summary>
        ''' YellowGreen Color:
        ''' Hex: "#FF9ACD32"
        ''' R=154, G=205, B=50
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property YellowGreen As Primitive = "#FF9ACD32"

        Public Shared Function GetHexaName(brush As System.Windows.Media.SolidColorBrush) As String
            Dim color = brush.Color
            Return $"#{Color.A:X2}{Color.R:X2}{Color.G:X2}{Color.B:X2}"
        End Function

        Public Shared Function GetHexaName(color As System.Windows.Media.Color) As String
            Return $"#{color.A:X2}{color.R:X2}{color.G:X2}{color.B:X2}"
        End Function


        ''' <summary>
        ''' Gets the color of active window's border as defined on the user system.
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property SystemActiveBorder As Primitive
            Get
                Return GetHexaName(SysColors.ActiveBorderBrush)
            End Get
        End Property

        ''' <summary>
        ''' Gets the color used on the usrs system to highlight a selected item that is inactive.
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property SystemInactiveSelectionHighlight As Primitive
            Get
                Return GetHexaName(SysColors.InactiveSelectionHighlightBrush)
            End Get
        End Property

        ''' <summary>
        ''' Gets the color of an inactive selected item’s text, as defined on the user system.
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property SystemInactiveSelectionHighlightText As Primitive
            Get
                Return GetHexaName(SysColors.InactiveSelectionHighlightTextBrush)
            End Get
        End Property

        ''' <summary>
        ''' Gets the background color in the client area of a window, as defined on the user system.
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property SystemWindow As Primitive
            Get
                Return GetHexaName(SysColors.WindowColor)
            End Get
        End Property

        ''' <summary>
        ''' Gets the background color of a scroll bar, as defined on the user system.
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property SystemScrollBar As Primitive
            Get
                Return GetHexaName(SysColors.ScrollBarColor)
            End Get
        End Property

        ''' <summary>
        ''' Gets the color of a menu's text, as defined on the user system.
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property SystemMenuText As Primitive
            Get
                Return GetHexaName(SysColors.MenuTextColor)
            End Get
        End Property

        ''' <summary>
        ''' Gets the color used to highlight a menu item, as defined on the user system.
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property SystemMenuHighlight As Primitive
            Get
                Return GetHexaName(SysColors.MenuHighlightColor)
            End Get
        End Property

        ''' <summary>
        ''' Gets the background color for a menu bar, as defined on the user system.
        ''' </summary>
        '     The background color for a menu bar, as defined on the user system.
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property SystemMenuBar As Primitive
            Get
                Return GetHexaName(SysColors.MenuBarColor)
            End Get
        End Property

        ''' <summary>
        ''' Gets the color of a menu's background, as defined on the user system.
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property SystemMenu As Primitive
            Get
                Return GetHexaName(SysColors.MenuColor)
            End Get
        End Property

        ''' <summary>
        ''' Gets the text color for the ToolTip control, as defined on the user system.
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property SystemInfoText As Primitive
            Get
                Return GetHexaName(SysColors.InfoTextColor)
            End Get
        End Property

        ''' <summary>
        ''' Gets the background color for the ToolTip control, as defined on the user system.
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property SystemInfo As Primitive
            Get
                Return GetHexaName(SysColors.InfoColor)
            End Get
        End Property

        ''' <summary>
        ''' Gets the color of the text of an inactive window's title bar, as defined on the user system.
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property SystemInactiveCaptionText As Primitive
            Get
                Return GetHexaName(SysColors.InactiveCaptionTextColor)
            End Get
        End Property

        ''' <summary>
        ''' Gets the background color of an inactive window's title bar, as defined on the user system.
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property SystemInactiveCaption As Primitive
            Get
                Return GetHexaName(SysColors.InactiveCaptionColor)
            End Get
        End Property

        ''' <summary>
        ''' Gets the color of an inactive window's border, as defined on the user system.
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property SystemInactiveBorder As Primitive
            Get
                Return GetHexaName(SysColors.InactiveBorderColor)
            End Get
        End Property

        ''' <summary>
        ''' Gets the color used to designate a hot-tracked item, as defined on the user system.
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property SystemHotTrack As Primitive
            Get
                Return GetHexaName(SysColors.HotTrackColor)
            End Get
        End Property

        ''' <summary>
        ''' Gets the color of the text of selected items, as defined on the user system.
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property SystemHighlightText As Primitive
            Get
                Return GetHexaName(SysColors.HighlightTextColor)
            End Get
        End Property

        ''' <summary>
        ''' Gets the background color of selected items, as defined on the user system.
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property SystemHighlight As Primitive
            Get
                Return GetHexaName(SysColors.HighlightColor)
            End Get
        End Property

        ''' <summary>
        ''' Gets the color of disabled text, as defined on the user system.
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property SystemGrayText As Primitive
            Get
                Return GetHexaName(SysColors.GrayTextColor)
            End Get
        End Property

        ''' <summary>
        ''' Gets the right side color in the gradient of an inactive window's title bar, as defined on the user system.
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property SystemGradientInactiveCaption As Primitive
            Get
                Return GetHexaName(SysColors.GradientInactiveCaptionColor)
            End Get
        End Property

        ''' <summary>
        ''' Gets the right side color in the gradient of an active window's title bar, as defined on the user system.
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property SystemGradientActiveCaption As Primitive
            Get
                Return GetHexaName(SysColors.GradientActiveCaptionColor)
            End Get
        End Property

        ''' <summary>
        ''' Gets the color of the desktop, as defined on the user system.
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property SystemDesktop As Primitive
            Get
                Return GetHexaName(SysColors.DesktopColor)
            End Get
        End Property

        ''' <summary>
        ''' Gets the color of text in a three-dimensional display element, as defined on the user system.
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property SystemControlText As Primitive
            Get
                Return GetHexaName(SysColors.ControlTextColor)
            End Get
        End Property

        ''' <summary>
        ''' Gets the highlight color of a three-dimensional display element, as defined on the user system.
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property SystemControlHighlight As Primitive
            Get
                Return GetHexaName(SysColors.ControlLightLightColor)
            End Get
        End Property

        ''' <summary>
        ''' Gets the light color of a three-dimensional display element, as defined on the user system.
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property SystemControlLight As Primitive
            Get
                Return GetHexaName(SysColors.ControlLightColor)
            End Get
        End Property

        ''' <summary>
        ''' Gets the dark shadow color of a three-dimensional display element, as defined on the user system.
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property SystemControlDarkShadow As Primitive
            Get
                Return GetHexaName(SysColors.ControlDarkDarkColor)
            End Get
        End Property

        ''' <summary>
        ''' Gets the shadow color of a three-dimensional display element, as defined on the user system.
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property SystemControlShadow As Primitive
            Get
                Return GetHexaName(SysColors.ControlDarkColor)
            End Get
        End Property

        ''' <summary>
        ''' Gets the face color of a three-dimensional display element, as defined on the user system.
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property SystemControl As Primitive
            Get
                Return GetHexaName(SysColors.ControlColor)
            End Get
        End Property

        ''' <summary>
        ''' Gets the color of the application workspace, as defined on the user system.
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property SystemAppWorkspace As Primitive
            Get
                Return GetHexaName(SysColors.AppWorkspaceColor)
            End Get
        End Property

        ''' <summary>
        ''' Gets the color of the text in      the active window's title bar, as defined on the user system.
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property SystemActiveCaptionText As Primitive
            Get
                Return GetHexaName(SysColors.ActiveCaptionTextColor)
            End Get
        End Property

        ''' <summary>
        ''' Gets the background color of the active window's title bar, as defined on the user system.
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property SystemActiveCaption As Primitive
            Get
                Return GetHexaName(SysColors.ActiveCaptionColor)
            End Get
        End Property

        ''' <summary>
        ''' Gets the color of a window frame, as defined on the user system.
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property SystemWindowFrame As Primitive
            Get
                Return GetHexaName(SysColors.WindowFrameColor)
            End Get
        End Property

        ''' <summary>
        ''' Gets the color of the text in the client area of a window, as defined on the user system.
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        Public Shared ReadOnly Property SystemWindowText As Primitive
            Get
                Return GetHexaName(SysColors.WindowTextColor)
            End Get
        End Property

    End Class
End Namespace
