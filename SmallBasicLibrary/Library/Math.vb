
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
        ''' <param name="number">
        ''' The number to get the absolute value for.
        ''' </param>
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
        ''' Gets the cosine of the given angle in radians.
        ''' </summary>
        ''' <param name="angle">
        ''' The angle whose cosine is needed (in radians).
        ''' </param>
        ''' <returns>
        ''' The cosine of the given angle.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Function Cos(angle As Primitive) As Primitive
            Return DoubleToDecimal(System.Math.Cos(angle))
        End Function

        ''' <summary>
        ''' Gets the sine of the given angle in radians.
        ''' </summary>
        ''' <param name="angle">
        ''' The angle whose sine is needed (in radians)
        ''' </param>
        ''' <returns>
        ''' The sine of the given angle
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Function Sin(angle As Primitive) As Primitive
            Return DoubleToDecimal(System.Math.Sin(angle))
        End Function

        ''' <summary>
        ''' Gets the tangent of the given angle in radians.
        ''' </summary>
        ''' <param name="angle">
        ''' The angle whose tangent is needed (in radians).
        ''' </param>
        ''' <returns>
        ''' The tangent of the given angle.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Function Tan(angle As Primitive) As Primitive
            Return DoubleToDecimal(System.Math.Tan(angle))
        End Function

        ''' <summary>
        ''' Gets the angle in radians, given the sine value.
        ''' </summary>
        ''' <param name="sinValue">
        ''' The sine value whose angle is needed.
        ''' </param>
        ''' <returns>
        ''' The angle (in radians) for the given sine Value.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Function ArcSin(sinValue As Primitive) As Primitive
            Return DoubleToDecimal(System.Math.Asin(sinValue))
        End Function

        ''' <summary>
        ''' Gets the angle in radians, given the cosine value.
        ''' </summary>
        ''' <param name="cosValue">
        ''' The cosine value whose angle is needed.
        ''' </param>
        ''' <returns>
        ''' The angle (in radians) for the given cosine Value.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Function ArcCos(cosValue As Primitive) As Primitive
            Return DoubleToDecimal(System.Math.Acos(cosValue))
        End Function

        ''' <summary>
        ''' Gets the angle in radians, given the tangent value.
        ''' </summary>
        ''' <param name="tanValue">
        ''' The tangent value whose angle is needed.
        ''' </param>
        ''' <returns>
        ''' The angle (in radians) for the given tangent Value.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Function ArcTan(tanValue As Primitive) As Primitive
            Return DoubleToDecimal(System.Math.Atan(tanValue))
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
            Return System.Math.Round(180.0 * angle.AsDecimal / System.Math.PI, 5) Mod 360.0
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
            Return DoubleToDecimal((angle.AsDecimal Mod 360.0) * System.Math.PI / 180.0)
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
        ''' <paramref name="decimalPlaces">The number of decimal places to keep in the number.</paramref>
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
        ''' </summary>
        ''' <param name="dividend">
        ''' The number to divide.
        ''' </param>
        ''' <param name="divisor">
        ''' The number that divides.
        ''' </param>
        ''' <returns>
        ''' The remainder after the division.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Function Remainder(dividend As Primitive, divisor As Primitive) As Primitive
            Return dividend.AsDecimal Mod divisor.AsDecimal
        End Function

        ''' <summary>
        ''' Gets a random number between 1 and the specified maxNumber (inclusive).
        ''' </summary>
        ''' <param name="maxNumber">
        ''' The maximum number for the requested random value.
        ''' </param>
        ''' <returns>
        ''' A Random number that is less than or equal to the specified max.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Function GetRandomNumber(maxNumber As Primitive) As Primitive
            If CBool((maxNumber < 1)) Then
                maxNumber = 1
            End If

            If _random Is Nothing Then
                _random = New Random(Now.Ticks Mod Integer.MaxValue)
            End If

            Return _random.Next(maxNumber) + 1
        End Function

        ''' <summary>
        ''' Handles double to decimal conversion
        ''' </summary>
        ''' <param name="number">
        ''' The input number
        ''' </param>
        ''' <returns>
        ''' The output number
        ''' </returns>
        Private Shared Function DoubleToDecimal(number As Double) As Primitive
            If Double.IsNaN(number) Then Return New Primitive(0)

            If Double.IsPositiveInfinity(number) Then
                Return New Primitive(Decimal.MaxValue)
            End If

            If Double.IsNegativeInfinity(number) Then
                Return New Primitive(Decimal.MinValue)
            End If

            Return New Primitive(number)
        End Function

        ''' <summary>
        ''' Converts the given decimal number to its hexadecimal representaion.
        ''' </summary>
        ''' <param name="decimal">The decimal number</param>
        ''' <returns>the hexadecimal representaion of the number</returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function Hex([decimal] As Primitive) As Primitive
            Return Conversion.Hex([decimal])
        End Function

    End Class
End Namespace
