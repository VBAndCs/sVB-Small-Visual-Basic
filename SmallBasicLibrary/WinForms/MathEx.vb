
Imports Microsoft.SmallVisualBasic.Library

Namespace WinForms

    <SmallVisualBasicType>
    <HideFromIntellisense>
    Public NotInheritable Class MathEx

        ''' <summary>
        ''' Gets the absolute value of the current number.  For example, -32.233 will return 32.233.
        ''' </summary>
        ''' <returns>The absolute value of the current number.</returns>
        <ReturnValueType(VariableType.Double)>
        <ExProperty>
        Public Shared Function GetAbs(number As Primitive) As Primitive
            Return Math.Abs(number)
        End Function

        ''' <summary>
        ''' Returns the smallest integer that is greater than or equal to the argument.  It rounds up the integer value.
        ''' For example, 32.233 will return 33. Also, 44 will return 44.
        ''' </summary>
        ''' <returns>
        ''' The ceiling value of the current number.
        ''' </returns>
        <ReturnValueType(VariableType.Double)>
        <ExProperty>
        Public Shared Function GetCeiling(number As Primitive) As Primitive
            Return Math.Ceiling(number)
        End Function

        ''' <summary>
        ''' Returns the largest integer that is less than or equal to the argument.  It rounds down the integer value.
        ''' For example, 32.233 will return 32. Also, 44 will return 44.
        ''' </summary>
        ''' <returns>
        ''' The floor value of the current number.
        ''' </returns>
        <ReturnValueType(VariableType.Double)>
        <ExProperty>
        Public Shared Function GetFloor(number As Primitive) As Primitive
            Return Math.Floor(number)
        End Function

        ''' <summary>
        ''' Gets the natural logarithm value of the current number.
        ''' </summary>
        ''' <returns>
        ''' The natural log value of the current number.
        ''' </returns>
        <ReturnValueType(VariableType.Double)>
        <ExProperty>
        Public Shared Function GetNaturalLog(number As Primitive) As Primitive
            Return Math.NaturalLog(number)
        End Function

        ''' <summary>
        ''' Gets the logarithm (base 10) value of the current number.
        ''' </summary>
        ''' <returns>
        ''' The log value of the current number
        ''' </returns>
        <ReturnValueType(VariableType.Double)>
        <ExProperty>
        Public Shared Function GetLog(number As Primitive) As Primitive
            Return Math.Log(number)
        End Function

        ''' <summary>
        ''' Gets the cosine of the current angle in radians.
        ''' </summary>
        ''' <returns>
        ''' The cosine of the current angle.
        ''' </returns>
        <ReturnValueType(VariableType.Double)>
        <ExProperty>
        Public Shared Function GetCos(angle As Primitive) As Primitive
            Return Math.Cos(angle)
        End Function

        ''' <summary>
        ''' Gets the sine of the current angle in radians.
        ''' </summary>
        ''' <returns>
        ''' The sine of the current angle
        ''' </returns>
        <ReturnValueType(VariableType.Double)>
        <ExProperty>
        Public Shared Function GetSin(angle As Primitive) As Primitive
            Return Math.Sin(angle)
        End Function

        ''' <summary>
        ''' Gets the tangent of the current angle in radians.
        ''' </summary>
        ''' <returns>
        ''' The tangent of the current angle.
        ''' </returns>
        <ReturnValueType(VariableType.Double)>
        <ExProperty>
        Public Shared Function GetTan(angle As Primitive) As Primitive
            Return Math.Tan(angle)
        End Function

        ''' <summary>
        ''' Gets the angle in radians for the current sin value.
        ''' </summary>
        ''' <returns>
        ''' The angle (in radians) for the current sine Value.
        ''' </returns>
        <ReturnValueType(VariableType.Double)>
        <ExProperty>
        Public Shared Function GetArcSin(sinValue As Primitive) As Primitive
            Return Math.ArcSin(sinValue)
        End Function

        ''' <summary>
        ''' Gets the angle in radians for the current cosine value.
        ''' </summary>
        ''' <returns>
        ''' The angle (in radians) for the current cosine Value.
        ''' </returns>
        <ReturnValueType(VariableType.Double)>
        <ExProperty>
        Public Shared Function GetArcCos(cosValue As Primitive) As Primitive
            Return Math.ArcCos(cosValue)
        End Function

        ''' <summary>
        ''' Gets the angle in radians for the current tangent value.
        ''' </summary>
        ''' <returns>
        ''' The angle (in radians) for the current tangent Value.
        ''' </returns>
        <ReturnValueType(VariableType.Double)>
        <ExProperty>
        Public Shared Function GetArcTan(tanValue As Primitive) As Primitive
            Return Math.ArcTan(tanValue)
        End Function

        ''' <summary>
        ''' Converts the current angle from radians to degrees.
        ''' </summary>
        ''' <returns>
        ''' The converted angle in degrees.
        ''' </returns>
        <ReturnValueType(VariableType.Double)>
        <ExProperty>
        Public Shared Function GetDegrees(angle As Primitive) As Primitive
            Return Math.GetDegrees(angle)
        End Function

        ''' <summary>
        ''' Converts the current angle from degrees to radians.
        ''' </summary>
        ''' <returns>
        ''' The converted angle in radians.
        ''' </returns>
        <ReturnValueType(VariableType.Double)>
        <ExProperty>
        Public Shared Function GetRadians(angle As Primitive) As Primitive
            Return Math.GetRadians(angle)
        End Function

        ''' <summary>
        ''' Gets the square root of the current number.
        ''' </summary>
        ''' <returns>
        ''' The square root value of the current number.
        ''' </returns>
        <ReturnValueType(VariableType.Double)>
        <ExProperty>
        Public Shared Function GetSquareRoot(number As Primitive) As Primitive
            Return Math.SquareRoot(number)
        End Function

        ''' <summary>
        ''' Raises the current number to the specified power.
        ''' </summary>
        ''' <param name="exponent">
        ''' The power to raise the base number.
        ''' </param>
        ''' <returns>
        ''' The base number raised to the specified exponent.
        ''' </returns>
        <ReturnValueType(VariableType.Double)>
        <ExMethod>
        Public Shared Function Power(baseNumber As Primitive, exponent As Primitive) As Primitive
            Return Math.Power(baseNumber, exponent)
        End Function

        ''' <summary>
        ''' Rounds the current number to the nearest integer.  For example 32.233 will be rounded to 32.0 while 
        ''' 32.566 will be rounded to 33.
        ''' </summary>
        ''' <returns>
        ''' The rounded value of the current number.
        ''' </returns>
        <ReturnValueType(VariableType.Double)>
        <ExProperty>
        Public Shared Function GetRound(number As Primitive) As Primitive
            Return Math.Round(number)
        End Function

        ''' <summary>
        ''' Rounds a current number to the given decimal places.
        ''' </summary>
        ''' <param name="decimalPlaces">the number of decimal places to keep in the number</param>
        ''' <returns>
        ''' The rounded value of the current number.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.Double)>
        <ExMethod>
        Public Shared Function Round(number As Primitive, decimalPlaces As Primitive) As Primitive
            Return System.Math.Round(number.AsDecimal, CInt(decimalPlaces.AsDecimal()), MidpointRounding.AwayFromZero)
        End Function

        ''' <summary>
        ''' Divides the first number by the second and returns the remainder.
        ''' Note that you can directly use the Mod operator to get the same result.
        ''' </summary>
        ''' <param name="divisor">
        ''' The number that divides. It can be positive or negative, but it can't be zero, otherwise this methodd will caause an error.
        ''' </param>
        ''' <returns>
        ''' The remainder after the division. It can be positive, negative or zero.
        ''' An example of a negative remainder is Remaibder(-10, -3) which returns -1, because
        ''' -10 = -3 * 3 - 1
        ''' </returns>
        <ExMethod>
        Public Shared Function Remainder(dividend As Primitive, divisor As Primitive) As Primitive
            Return Math.Remainder(dividend, divisor)
        End Function

        ''' <summary>
        ''' Gets a random number between 1 and the specified maxNumber (inclusive).
        ''' </summary>
        ''' <returns>
        ''' A Random number that is less than or equal to the specified max.
        ''' </returns>
        <ReturnValueType(VariableType.Double)>
        <ExProperty>
        Public Shared Function GetRandom(maxNumber As Primitive) As Primitive
            Return Math.GetRandomNumber(maxNumber)
        End Function


        ''' <summary>
        ''' Creates a new date from the current ticks value. Note that the second contains 10 milion ticks.
        ''' </summary>
        ''' <returns>a new date.</returns>
        <ReturnValueType(VariableType.Date)>
        <ExMethod>
        Public Shared Function ToDate(ticks As Primitive) As Primitive
            Return New Primitive(ticks.AsDecimal(), NumberType.Date)
        End Function

        ''' <summary>
        ''' Creates a new TimeSpan from the current ticks value. Note that the second contains 10 milion ticks.
        ''' </summary>
        ''' <returns>a new duration</returns>
        <ReturnValueType(VariableType.Date)>
        <ExMethod>
        Public Shared Function ToDuration(ticks As Primitive) As Primitive
            Return New Primitive(ticks.AsDecimal, NumberType.TimeSpan)
        End Function


        ''' <summary>
        ''' Gets the character that crossponds to the current Unicode value.
        ''' </summary>
        ''' <returns>the character of the current Unicode value</returns>
        <ReturnValueType(VariableType.String)>
        <ExMethod>
        Public Shared Function ToChar(unicode As Primitive) As Primitive
            Return Chars.GetCharacter(unicode)
        End Function

        ''' <summary>
        ''' Converts the current decimal number to its hexadecimal representaion.
        ''' </summary>
        ''' <returns>the hexadecimal representaion of the number</returns>
        <WinForms.ReturnValueType(VariableType.String)>
        <ExProperty>
        Public Shared Function GetHex([decimal] As Primitive) As Primitive
            Return New Primitive(Hex([decimal]))
        End Function

        ''' <summary>
        ''' Converts the current hexadecimal number to a decimal number. 
        ''' You can use the Math.Hex method to convert a decimal number to a hexadecimal number.
        ''' </summary>
        ''' <returns>The decimal value the hexa number if it is valid, or an empty string otherwise.</returns>
        <WinForms.ReturnValueType(VariableType.Double)>
        <ExMethod>
        Public Shared Function ToDecimal(hex As Primitive) As Primitive
            Return System.Convert.ToInt32(hex, 16)
        End Function


        ''' <summary>
        ''' Gets the integeral part of the number, which means that the decimal part will be renoved without doing any rounding.
        ''' For example, 32.233 will return 32 and 44.7 will return 44.
        ''' </summary>
        ''' <returns>
        ''' The integral part of the number.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.Double)>
        <ExMethod>
        Public Shared Function Truncate(number As Primitive) As Primitive
            Return System.Math.Truncate(number.AsDecimal)
        End Function

        ''' <summary>
        ''' Limits the current number to be in the range of the given min and max values.
        ''' </summary>
        ''' <param name="minValue">The minimun value that the number should not be less than.</param>
        ''' <param name="maxValue">The maximun value that the number should not be greater than.</param>
        ''' <returns>
        ''' The minValue if the number is less than,
        ''' or the maxValue is the number is greater than,
        ''' otherwise returns the number itself.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.Double)>
        <ExMethod>
        Public Shared Function Limit(number As Primitive, minValue As Primitive, maxValue As Primitive) As Primitive
            Return Math.Limit(number, minValue, maxValue)
        End Function
    End Class
End Namespace
