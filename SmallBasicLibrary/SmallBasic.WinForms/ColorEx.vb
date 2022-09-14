Imports System.Windows.Media
Imports Microsoft.SmallBasic.Library

Namespace WinForms

    <SmallBasicType>
    <HideFromIntellisense>
    Public NotInheritable Class ColorEx

        ''' <summary>
        ''' Crates a new color based on the current color, with the transparency set to the given value.
        ''' </summary>
        ''' <param name="percentage">a 0 to 100 value that represents the percentage of the transparency of the color</param>
        ''' <returns>a new color with the given transparency. The current color will not change.</returns>
        <ReturnValueType(VariableType.Color)>
        <ExMethod>
        Public Shared Function ChangeTransparency(color As Primitive, percentage As Primitive) As Primitive
            Return WinForms.Color.ChangeTransparency(color, percentage)
        End Function

        ''' <summary>
        ''' reterns the transparency percentage of the color
        ''' </summary>
        ''' <returns>a 0 to 100 value represents the transparency percentage of the color</returns>
        <ReturnValueType(VariableType.Double)>
        <WinForms.ExProperty>
        Public Shared Function GetTransparency(color As Primitive) As Primitive
            Return WinForms.Color.GetTransparency(color)
        End Function

        ''' <summary>
        ''' Returms the English name name of the color if its defined.
        ''' </summary>
        ''' <returns>the English name name of the color</returns>
        <ReturnValueType(VariableType.String)>
        <WinForms.ExProperty> Public Shared Function GetName(color As Primitive) As Primitive
            Return WinForms.Color.GetName(color)
        End Function

        ''' <summary>
        ''' Returms the English name name of the color if its defined, followed by the transparency percentage of the color.
        ''' </summary>
        ''' <returns>the English name name of the color, followed by the transparency percentage of the color</returns>
        <WinForms.ExProperty>
        <ReturnValueType(VariableType.String)>
        Public Shared Function GetNameAndTransparency(color As Primitive) As Primitive
            Return WinForms.Color.GetNameAndTransparency(color)
        End Function


        ''' <summary>
        ''' Gets the red component of the color
        ''' </summary>
        ''' <returns>A number betwwen 0 and 255 thet represents the red ration in the color</returns>
        <ReturnValueType(VariableType.Double)>
        <ExProperty>
        Public Shared Function GetRedRatio(color As Primitive) As Primitive
            Return WinForms.Color.GetRedRatio(color)
        End Function

        ''' <summary>
        ''' Gets the green component of the color
        ''' </summary>
        ''' <returns>A number betwwen 0 and 255 thet represents the green ration in the color</returns>
        <ReturnValueType(VariableType.Double)>
        <ExProperty>
        Public Shared Function GetGreenRatio(color As Primitive) As Primitive
            Return WinForms.Color.GetGreenRatio(color)
        End Function

        ''' <summary>
        ''' Gets the blue component of the color
        ''' </summary>
        ''' <returns>A number betwwen 0 and 255 thet represents the blue ration in the color</returns>

        <ReturnValueType(VariableType.Double)>
        <ExProperty>
        Public Shared Function GetBlueRatio(color As Primitive) As Primitive
            Return WinForms.Color.GetBlueRatio(color)
        End Function

        ''' <summary>
        ''' Creates a new color based on the current color, with the red component changed to the given value.
        ''' </summary>
        ''' <param name="value">the new value of the red component</param>
        ''' <returns>a new color with the red component changed to the given value. The current color will not change.</returns>
        <ReturnValueType(VariableType.Color)>
        <ExMethod>
        Public Shared Function ChangeRedRatio(color As Primitive, value As Primitive) As Primitive
            Return WinForms.Color.ChangeRedRatio(color, value)
        End Function

        ''' <summary>
        ''' Creates a new color based on the current color, with the green component changed to the given value.
        ''' </summary>
        ''' <param name="value">the new value of the green component</param>
        ''' <returns>a new color with the green component changed to the given value. The current color will not change.</returns>
        <ReturnValueType(VariableType.Color)>
        <ExMethod>
        Public Shared Function ChangeGreenRatio(color As Primitive, value As Primitive) As Primitive
            Return WinForms.Color.ChangeGreenRatio(color, value)
        End Function

        ''' <summary>
        ''' Creates a new color based on the current color, with the blue component changed to the given value.
        ''' </summary>
        ''' <param name="value">the new value of the blue component</param>
        ''' <returns>a new color with the blue component changed to the given value. The current color will not change.</returns>
        <ReturnValueType(VariableType.Color)>
        <ExMethod>
        Public Shared Function ChangeBlueRatio(color As Primitive, value As Primitive) As Primitive
            Return WinForms.Color.ChangeBlueRatio(color, value)
        End Function

    End Class
End Namespace