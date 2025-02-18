
Namespace Library
    ''' <summary>
    ''' The Math class provides lots of useful mathematics related methods
    ''' </summary>
    <SmallVisualBasicType>
    Public NotInheritable Class Math
        Private Shared _random As Random

        ''' <summary>
        ''' Gets the value of Pi, which equals 3.1415926535897931
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared ReadOnly Property Pi As Primitive = System.Math.PI

        ''' <summary>
        ''' Gets the absolute value of the given number.  For example, -32.233 will return 32.233.
        ''' </summary>
        ''' <param name="number">The number to get the absolute value for.</param>
        ''' <returns>
        ''' The absolute value of the given number.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Function Abs(number As Primitive) As Primitive
            Return System.Math.Abs(number.AsDecimal)
        End Function

        ''' <summary>
        ''' Returns the smallest integer that is greater than or equal to the argument.  It rounds up the integer value.
        ''' For example, 32.233 will return 33. Also, 44 will return 44.
        ''' </summary>
        ''' <param name="number">
        ''' The number whose ceiling is required.
        ''' </param>
        ''' <returns>
        ''' The ceiling value of the given number.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Function Ceiling(number As Primitive) As Primitive
            Return System.Math.Ceiling(number.AsDecimal)
        End Function

        ''' <summary>
        ''' Returns the largest integer that is less than or equal to the argument.  It rounds down the integer value.
        ''' For example, 32.233 will return 32. Also, 44 will return 44.
        ''' </summary>
        ''' <param name="number">
        ''' The number whose floor value is required.
        ''' </param>
        ''' <returns>
        ''' The floor value of the given number.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Function Floor(number As Primitive) As Primitive
            Return System.Math.Floor(number.AsDecimal)
        End Function

        ''' <summary>
        ''' Gets the integeral part of the given number, which means that the decimal part will be renoved without doing any rounding.
        ''' For example, 32.233 will return 32 and 44.7 will return 44.
        ''' </summary>
        ''' <param name="number">
        ''' The number whose truncated value is required.
        ''' </param>
        ''' <returns>
        ''' The integral part of the given number.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Function Truncate(number As Primitive) As Primitive
            Return System.Math.Truncate(number.AsDecimal)
        End Function

        ''' <summary>
        ''' Gets the natural logarithm value of the given number.
        ''' </summary>
        ''' <param name="number">The number whose natural logarithm value is required.</param>
        ''' <returns>
        ''' The natural log value of the given number.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Function NaturalLog(number As Primitive) As Primitive
            Return DoubleToDecimal(System.Math.Log(number))
        End Function

        ''' <summary>
        ''' Gets the natural logarithmic base, which equals 2.7182818284590451
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared ReadOnly Property E As Primitive = System.Math.E

        ''' <summary>
        ''' Gets the number resulted from raisng the natural logarithmic base (E) to the given power
        ''' </summary>
        ''' <param name="exponent">The number to be used as the power of E.</param>
        ''' <returns>E raised to the specified exponent.</returns>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Function Exp(exponent As Primitive) As Primitive
            Return New Primitive(System.Math.Exp(exponent))
        End Function

        ''' <summary>
        ''' Gets the logarithm (base 10) value of the given number.
        ''' </summary>
        ''' <param name="number">
        ''' The number whose logarithm value is required
        ''' </param>
        ''' <returns>
        ''' The log value of the given number
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Function Log(number As Primitive) As Primitive
            Return DoubleToDecimal(System.Math.Log10(number))
        End Function

        ''' <summary>
        ''' Indicates whither or not to consider angle values to be in radian when sent to Sin, Cos and Tan or returned from ArcSin, ArcCos, and ArcTan.
        ''' The default value is True, but you can set if to False to use degress instead.
        ''' Note that this property will also affect how the Evaluator works.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Boolean)>
        Public Shared Property UseRadianAngles As Primitive = True

        ''' <summary>
        ''' Gets the cosine of the given angle in radians.
        ''' </summary>
        ''' <param name="angle">
        ''' The angle whose cosine is needed (in radians).
        ''' If you have the angle in degrees, use the GetRadians method to convert it to reeadians, or set the UseRadianAngles property fo False to force all trigonometric functions to work with angles in degrees.        ''' </param>
        ''' <returns>
        ''' The cosine of the given angle.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Function Cos(angle As Primitive) As Primitive
            If UseRadianAngles Then
                Return DoubleToDecimal(System.Math.Cos(angle))
            Else
                Return DoubleToDecimal(System.Math.Cos(CDbl(angle) * System.Math.PI / 180.0))
            End If
        End Function

        ''' <summary>
        ''' Gets the sine of the given angle in radians.
        ''' </summary>
        ''' <param name="angle">
        ''' The angle whose sine is needed (in radians)
        ''' If you have the angle in degrees, use the GetRadians method to convert it to reeadians, or set the UseRadianAngles property fo False to force all trigonometric functions to work with angles in degrees.
        ''' </param>
        ''' <returns>
        ''' The sine of the given angle
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Function Sin(angle As Primitive) As Primitive
            If UseRadianAngles Then
                Return DoubleToDecimal(System.Math.Sin(angle))
            Else
                Return DoubleToDecimal(System.Math.Sin(CDbl(angle) * System.Math.PI / 180.0))
            End If
        End Function

        ''' <summary>
        ''' Gets the tangent of the given angle in radians.
        ''' </summary>
        ''' <param name="angle">
        ''' The angle whose tangent is needed (in radians).
        ''' If you have the angle in degrees, use the GetRadians method to convert it to reeadians, or set the UseRadianAngles property fo False to force all trigonometric functions to work with angles in degrees.        ''' </param>
        ''' <returns>
        ''' The tangent of the given angle.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Function Tan(angle As Primitive) As Primitive
            If UseRadianAngles Then
                Return DoubleToDecimal(System.Math.Tan(angle))
            Else
                Return DoubleToDecimal(System.Math.Tan(CDbl(angle) * System.Math.PI / 180.0))
            End If
        End Function

        ''' <summary>
        ''' Gets the angle in radians, given the sine value.
        ''' </summary>
        ''' <param name="sinValue">
        ''' The sine value whose angle is needed.
        ''' </param>
        ''' <returns>
        ''' The angle (in radians) for the given sine Value.
        ''' Use the GetDegrees method to convert the angle to degrees, or set the UseRadianAngles property fo False to force all trigonometric functions to work with angles in degrees.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Function ArcSin(sinValue As Primitive) As Primitive
            If UseRadianAngles Then
                Return DoubleToDecimal(System.Math.Asin(sinValue))
            Else
                Return DoubleToDecimal(System.Math.Asin(sinValue) * 180.0 / System.Math.PI)
            End If
        End Function

        ''' <summary>
        ''' Gets the angle in radians, given the cosine value.
        ''' </summary>
        ''' <param name="cosValue">
        ''' The cosine value whose angle is needed.
        ''' </param>
        ''' <returns>
        ''' The angle (in radians) for the given cosine Value.
        ''' Use the GetDegrees method to convert the angle to degrees, or set the UseRadianAngles property fo False to force all trigonometric functions to work with angles in degrees.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Function ArcCos(cosValue As Primitive) As Primitive
            If UseRadianAngles Then
                Return DoubleToDecimal(System.Math.Acos(cosValue))
            Else
                Return DoubleToDecimal(System.Math.Acos(cosValue) * 180.0 / System.Math.PI)
            End If
        End Function

        ''' <summary>
        ''' Gets the angle in radians, given the tangent value.
        ''' </summary>
        ''' <param name="tanValue">
        ''' The tangent value whose angle is needed.
        ''' </param>
        ''' <returns>
        ''' The angle (in radians) for the given tangent Value.
        ''' Use the GetDegrees method to convert the angle to degrees, or set the UseRadianAngles property fo False to force all trigonometric functions to work with angles in degrees.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Function ArcTan(tanValue As Primitive) As Primitive
            If UseRadianAngles Then
                Return DoubleToDecimal(System.Math.Atan(tanValue))
            Else
                Return DoubleToDecimal(System.Math.Atan(tanValue) * 180.0 / System.Math.PI)
            End If
        End Function

        ''' <summary>
        ''' Converts a given angle in radians to degrees.
        ''' </summary>
        ''' <param name="angle">
        ''' The angle in radians.
        ''' </param>
        ''' <returns>
        ''' The converted angle in degrees.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Function GetDegrees(angle As Primitive) As Primitive
            Return Decimal.Remainder(System.Math.Round(180D * angle.AsDecimal / System.Math.PI, 5), 360D)
        End Function

        ''' <summary>
        ''' Converts a given angle in degrees to radians.
        ''' </summary>
        ''' <param name="angle">
        ''' The angle in degrees.
        ''' </param>
        ''' <returns>
        ''' The converted angle in radians.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Function GetRadians(angle As Primitive) As Primitive
            Return DoubleToDecimal((angle.AsDecimal Mod 360D) * System.Math.PI / 180.0)
        End Function

        ''' <summary>
        ''' Gets the square root of a given number.
        ''' </summary>
        ''' <param name="number">
        ''' The number whose square root value is needed.
        ''' </param>
        ''' <returns>
        ''' The square root value of the given number.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Function SquareRoot(number As Primitive) As Primitive
            If Not number.IsNumber OrElse number < 0 Then Return New Primitive("")
            Return DoubleToDecimal(System.Math.Sqrt(number))
        End Function

        ''' <summary>
        ''' Raises the base number to the specified power.
        ''' </summary>
        ''' <param name="baseNumber">
        ''' The number to be raised to the exponent power.
        ''' </param>
        ''' <param name="exponent">
        ''' The power to raise the base number.
        ''' </param>
        ''' <returns>
        ''' The base number raised to the specified exponent.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Function Power(baseNumber As Primitive, exponent As Primitive) As Primitive
            Return DoubleToDecimal(System.Math.Pow(baseNumber, exponent))
        End Function

        ''' <summary>
        ''' Rounds a given number to the nearest integer.  For example 32.233 will be rounded to 32.0 while 
        ''' 32.566 will be rounded to 33.
        ''' </summary>
        ''' <param name="number">
        ''' The number whose approximation is required.
        ''' </param>
        ''' <returns>
        ''' The rounded value of the given number.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Function Round(number As Primitive) As Primitive
            Return System.Math.Round(number.AsDecimal, MidpointRounding.AwayFromZero)
        End Function

        ''' <summary>
        ''' Rounds a given number to the given decimal places.
        ''' </summary>
        ''' <param name="number">The number whose approximation is required.</param>
        ''' <param name="decimalPlaces">The number of decimal places to keep in the number.</param>
        ''' <returns>
        ''' The rounded value of the given number.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Function Round2(number As Primitive, decimalPlaces As Primitive) As Primitive
            Return System.Math.Round(number.AsDecimal, CInt(decimalPlaces.AsDecimal()), MidpointRounding.AwayFromZero)
        End Function

        ''' <summary>
        ''' Compares two numbers and returns the greater of the two.
        ''' </summary>
        ''' <param name="number1">
        ''' The first of the two numbers to compare.
        ''' </param>
        ''' <param name="number2">
        ''' The second of the two numbers to compare.
        ''' </param>
        ''' <returns>
        ''' The greater value of the two numbers.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Function Max(number1 As Primitive, number2 As Primitive) As Primitive
            Return System.Math.Max(number1.AsDecimal, number2.AsDecimal)
        End Function

        ''' <summary>
        ''' Compares two numbers and returns the smaller of the two.
        ''' </summary>
        ''' <param name="number1">
        ''' The first of the two numbers to compare.
        ''' </param>
        ''' <param name="number2">
        ''' The second of the two numbers to compare.
        ''' </param>
        ''' <returns>
        ''' The smaller value of the two numbers.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Function Min(number1 As Primitive, number2 As Primitive) As Primitive
            Return System.Math.Min(number1.AsDecimal, number2.AsDecimal)
        End Function

        ''' <summary>
        ''' Divides the first number by the second and returns the remainder.
        ''' Note that you can directly use the Mod operator to get the same result.
        ''' </summary>
        ''' <param name="dividend">
        ''' The number to divide. This can be positive, negative or zero.
        ''' </param>
        ''' <param name="divisor">
        ''' The number that divides. It can be positive or negative, but it can't be zero, otherwise this methodd will caause an error.
        ''' </param>
        ''' <returns>
        ''' The remainder after the division. It can be positive, negative or zero.
        ''' An example of a negative remainder is Remaibder(-10, -3) which returns -1, because
        ''' -10 = -3 * 3 - 1
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Function Remainder(dividend As Primitive, divisor As Primitive) As Primitive
            Dim d = divisor.AsDecimal
            If d = 0 Then Return dividend
            Return Decimal.Remainder(dividend.AsDecimal, d)
        End Function

        ''' <summary>
        ''' Gets a random number between 1 and the specified maxNumber (inclusive).
        ''' </summary>
        ''' <param name="maxNumber">The maximum value for the requested random number.</param>
        ''' <returns>
        ''' A positive random number that is less than or equal to the specified max.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Function GetRandomNumber(maxNumber As Primitive) As Primitive
            If CBool((maxNumber < 1)) Then maxNumber = 1
            If _random Is Nothing Then _random = New Random()
            Return _random.Next(maxNumber) + 1
        End Function


        ''' <summary>
        ''' Calculates the distance between the given two points in the 2D space.
        ''' </summary>
        ''' <param name="x1">The x co-ordinate of the first point.</param>
        ''' <param name="y1">The y co-ordinate of the first point.</param>
        ''' <param name="x2">The x co-ordinate of the second point.</param>
        ''' <param name="y2">The y co-ordinate of the second point.</param>
        ''' ''' <returns>
        ''' The distance berween the two points.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Function GetDistance(
                   x1 As Primitive,
                   y1 As Primitive,
                   x2 As Primitive,
                   y2 As Primitive) As Primitive

            Return New Primitive(System.Math.Sqrt(
                CDbl(x2.AsDecimal() - x1.AsDecimal()) ^ 2.0 +
                CDbl(y2.AsDecimal() - y1.AsDecimal()) ^ 2.0
            ))
        End Function

        ''' <summary>
        ''' Calculates the distance between the given two points in the 3D space.
        ''' </summary>
        ''' <param name="x1">The x co-ordinate of the first point.</param>
        ''' <param name="y1">The y co-ordinate of the first point.</param>
        ''' <param name="z1">The z co-ordinate of the first point.</param>
        ''' <param name="x2">The x co-ordinate of the second point.</param>
        ''' <param name="y2">The y co-ordinate of the second point.</param>
        ''' <param name="z2">The z co-ordinate of the second point.</param>
        ''' ''' <returns>
        ''' The distance berween the two points.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Function GetDistance3D(
                   x1 As Primitive,
                   y1 As Primitive,
                   z1 As Primitive,
                   x2 As Primitive,
                   y2 As Primitive,
                   z2 As Primitive) As Primitive

            Return New Primitive(System.Math.Sqrt(
                CDbl(x2.AsDecimal() - x1.AsDecimal()) ^ 2.0 +
                CDbl(y2.AsDecimal() - y1.AsDecimal()) ^ 2.0 +
                CDbl(z2.AsDecimal() - z1.AsDecimal()) ^ 2.0
            ))
        End Function

        ''' <summary>
        ''' Gets a random number between the given two numbers (inclusive).
        ''' </summary>
        ''' <param name="minNumber">The minimum value for the requested random number.</param>
        ''' <param name="maxNumber">The maximum value for the requested random number.</param>
        ''' <returns>
        ''' A random number that is greater than or equal the specified min value, and less than or equal to the specified max value.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Function Rnd(minNumber As Primitive, maxNumber As Primitive) As Primitive
            Dim max, min As Integer
            If CBool((maxNumber < minNumber)) Then
                max = minNumber
                min = maxNumber
            Else
                min = minNumber
                max = maxNumber
            End If

            If max = min Then Return min

            If _random Is Nothing Then _random = New Random()
            Return min + _random.Next(max - min + 1)
        End Function

        ''' <summary>
        ''' Handles double to decimal conversion
        ''' </summary>
        ''' <param name="number">The input number</param>
        ''' <returns>The output number</returns>
        Private Shared Function DoubleToDecimal(number As Double) As Primitive
            If Double.IsNaN(number) Then
                Return New Primitive(0)
            ElseIf Double.IsPositiveInfinity(number) Then
                Return New Primitive(Decimal.MaxValue)
            ElseIf Double.IsNegativeInfinity(number) Then
                Return New Primitive(Decimal.MinValue)
            Else
                Return New Primitive(System.Math.Round(number, 10))
            End If
        End Function

        ''' <summary>
        ''' Converts the given decimal number to its hexadecimal representaion.
        ''' </summary>
        ''' <param name="decimal">The decimal number whose hexadecimal value is required.</param>
        ''' <returns>A string that represnts the hexadecimal value the number.</returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function Hex([decimal] As Primitive) As Primitive
            Return New Primitive(Conversion.Hex([decimal]))
        End Function

        ''' <summary>
        ''' Converts the given hexadecimal number to a decimal number. 
        ''' You can use the Math.Hex method to convert a decimal number to a hexadecimal number.
        ''' </summary>
        ''' <param name="hex">The hexadecimal number whose decimal value is required. This argument can be omitted if you call this method as an extension method (with the name ToDeciaml) of a numeric or string variables, because hexa numbrs can be a pure digital numbers like 10 which is the hexa represntaion of 16 in the decimal system, or it can contain letters from A to F which makes it a string.</param>
        ''' <returns>The decimal value the hexa number if it is valid, or an empty string otherwise.</returns>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Function [Decimal](hex As Primitive) As Primitive
            Return Convert.ToInt32(hex, 16)
        End Function

    End Class
End Namespace
