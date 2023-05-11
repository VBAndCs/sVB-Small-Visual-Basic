Imports System.Reflection

Imports System.Text

Friend NotInheritable Class ColorHelper


    Private Shared hexArray() As Char = {"0"c, "1"c, "2"c, "3"c, "4"c, "5"c, "6"c, "7"c, "8"c, "9"c, "A"c, "B"c, "C"c, "D"c, "E"c, "F"c}

    Public Shared Function MakeValidColorString(ByVal S As String) As String
        Dim _s As String = S

        For i As Integer = 0 To _s.Length - 1
            Dim c As Char = _s.Chars(i)

            If Not (c >= "a"c AndAlso c <= "f"c) AndAlso Not (c >= "A"c AndAlso c <= "F"c) AndAlso Not (c >= "0"c AndAlso c <= "9"c) Then
                _s = _s.Remove(i, 1)
                i -= 1
            End If
        Next i

        If _s.Length > 8 Then
            _s = _s.Substring(0, 8)
        End If

        Do While _s.Length <= 8 AndAlso _s.Length <> 3 AndAlso _s.Length <> 4 AndAlso _s.Length <> 6 AndAlso _s.Length <> 8
            _s = _s & "0"
        Loop

        Return _s
    End Function

    Public Shared Function ColorFromString(S As String) As Color
        Dim _s As String = MakeValidColorString(S)

        Dim A As Byte = 255
        Dim R As Byte = 0
        Dim G As Byte = 0
        Dim B As Byte = 0

        If _s.Length = 3 Then
            R = Byte.Parse(_s.Substring(0, 1) & _s.Substring(0, 1), System.Globalization.NumberStyles.HexNumber)
            G = Byte.Parse(_s.Substring(1, 1) & _s.Substring(1, 1), System.Globalization.NumberStyles.HexNumber)
            B = Byte.Parse(_s.Substring(2, 1) & _s.Substring(2, 1), System.Globalization.NumberStyles.HexNumber)
        End If

        If _s.Length = 4 Then
            A = Byte.Parse(_s.Substring(0, 1) & _s.Substring(0, 1), System.Globalization.NumberStyles.HexNumber)
            R = Byte.Parse(_s.Substring(1, 1) & _s.Substring(1, 1), System.Globalization.NumberStyles.HexNumber)
            G = Byte.Parse(_s.Substring(2, 1) & _s.Substring(2, 1), System.Globalization.NumberStyles.HexNumber)
            B = Byte.Parse(_s.Substring(3, 1) & _s.Substring(3, 1), System.Globalization.NumberStyles.HexNumber)
        End If

        If _s.Length = 6 Then
            R = Byte.Parse(_s.Substring(0, 2), Globalization.NumberStyles.HexNumber)
            G = Byte.Parse(_s.Substring(1, 2), Globalization.NumberStyles.HexNumber)
            B = Byte.Parse(_s.Substring(2, 2), Globalization.NumberStyles.HexNumber)
        End If

        If _s.Length = 8 Then
            A = Byte.Parse(_s.Substring(0, 2), Globalization.NumberStyles.HexNumber)
            R = Byte.Parse(_s.Substring(1, 2), Globalization.NumberStyles.HexNumber)
            G = Byte.Parse(_s.Substring(2, 2), Globalization.NumberStyles.HexNumber)
            B = Byte.Parse(_s.Substring(3, 2), Globalization.NumberStyles.HexNumber)
        End If

        Return Color.FromArgb(A, R, G, B)
    End Function

    Public Shared Function StringFromColor(c As Color) As String
        Dim bytes() As Byte = {c.A, c.R, c.G, c.B}

        Dim chars(bytes.Length * 2 - 1) As Char

        For i As Integer = 0 To bytes.Length - 1
            Dim b As Integer = bytes(i)
            chars(i * 2) = hexArray(b >> 4)
            chars(i * 2 + 1) = hexArray(b And &HF)
        Next i

        Return New String(chars)
    End Function

    Public Shared Function ColorFromHSB(H As Double, S As Double, ByVal B As Double) As Color
        Dim red As Double = 0.0, green As Double = 0.0, blue As Double = 0.0

        If S = 0.0 Then
            blue = B
            green = blue
            red = green
        Else
            Dim _h As Double = H * 360
            Do While _h >= 360.0
                _h -= 360.0
            Loop

            _h = _h / 60.0
            Dim i As Integer = CInt(Fix(_h))

            Dim f As Double = _h - i
            Dim r As Double = B * (1.0 - S)
            Dim _s As Double = B * (1.0 - S * f)
            Dim t As Double = B * (1.0 - S * (1.0 - f))

            Select Case i
                Case 0
                    red = B
                    green = t
                    blue = r
                Case 1
                    red = _s
                    green = B
                    blue = r
                Case 2
                    red = r
                    green = B
                    blue = t
                Case 3
                    red = r
                    green = _s
                    blue = B
                Case 4
                    red = t
                    green = r
                    blue = B
                Case 5
                    red = B
                    green = r
                    blue = _s
            End Select
        End If

        Dim iRed As Byte = CByte(red * 255.0)
        Dim iGreen As Byte = CByte(green * 255.0)
        Dim iBlue As Byte = CByte(blue * 255.0)

        Return Color.FromRgb(iRed, iGreen, iBlue)
    End Function

    Public Shared Sub HSBFromColor(ByVal C As Color, ByRef H As Double, ByRef S As Double, ByRef B As Double)
        Dim red As Byte = C.R
        Dim green As Byte = C.G
        Dim blue As Byte = C.B

        Dim imax As Integer = red, imin As Integer = red

        If green > imax Then
            imax = green
        ElseIf green < imin Then
            imin = green
        End If
        If blue > imax Then
            imax = blue
        ElseIf blue < imin Then
            imin = blue
        End If
        Dim max As Double = imax / 255.0, min As Double = imin / 255.0

        Dim value As Double = max
        Dim saturation As Double = If(max > 0, (max - min) / max, 0.0)
        Dim hue As Double = 0

        If imax > imin Then
            Dim f As Double = 1.0 / ((max - min) * 255.0)
            hue = If(imax = red, 0.0 + f * (CDbl(green) - blue), If(imax = green, 2.0 + f * (CDbl(blue) - red), 4.0 + f * (CDbl(red) - green)))
            hue = hue * 60.0
            If hue < 0.0 Then
                hue += 360.0
            End If
        End If

        H = hue / 360
        S = saturation
        B = value
    End Sub

    Public Shared Function ColorFromAHSB(ByVal A As Double, ByVal H As Double, ByVal S As Double, ByVal B As Double) As Color
        Dim r As Color = ColorFromHSB(H, S, B)
        r.A = CByte(Math.Round(A * 255))
        Return r
    End Function

    Public Shared Function GetColorName(ByVal color As Color) As String
        Return _knownColors.Where(Function(kvp) kvp.Value.Equals(color)).Select(Function(kvp) kvp.Key).FirstOrDefault()
    End Function

    Private Shared ReadOnly _knownColors As Dictionary(Of String, Color) = GetKnownColors()

    Private Shared Function GetKnownColors() As Dictionary(Of String, Color)
        Dim colorProperties = GetType(Colors).GetProperties(BindingFlags.Static Or BindingFlags.Public)
        Return colorProperties.ToDictionary(Function(p) p.Name, Function(p) CType(p.GetValue(Nothing, Nothing), Color))
    End Function

End Class
