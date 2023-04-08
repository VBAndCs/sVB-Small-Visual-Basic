Imports System.Globalization

Namespace Library
    ''' <summary>
    ''' The Text object provides helpful operations for working with Text.
    ''' </summary>
    <SmallVisualBasicType>
    Public NotInheritable Class Text

        ''' <summary>
        ''' Formats the string by replacing [1], [2], ... [n] by items from the values array.
        ''' </summary>
        ''' <param name="text">The string to Format. Use [1], [2],... [n] in the string, to refer the values[1], values[2], ... values[n]</param>
        ''' <param name="values">An array its elements will be used to replace [1], [2],... [n] strings if found in the text</param>
        ''' <returns>The formated string after substituting [1], [2],... [n] with elements from the values array</returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function Format(text As Primitive, values As Primitive) As Primitive
            Dim str = CStr(text)
            If str.Trim() = "" Then Return str

            Dim arr() As Primitive
            If values.IsArray Then
                arr = values._arrayMap.Values.ToArray()
            Else
                arr = {values}
            End If

            Dim sb As New System.Text.StringBuilder(str)
            For i = 1 To arr.Count
                sb.Replace($"[{i}]", $"<<[[{i}]]>>")
            Next

            For i = 0 To arr.Count - 1
                sb.Replace($"<<[[{i + 1}]]>>", arr(i))
            Next

            Return sb.ToString()
        End Function


        ''' <summary>
        ''' Checks if the string contains a numeric value
        ''' </summary>
        ''' <param name="text">the string to check its value</param>
        ''' <returns>True if text is a number, False otherwise</returns>
        <WinForms.ReturnValueType(VariableType.Boolean)>
        Public Shared Function IsNumeric(text As Primitive) As Primitive
            Return VisualBasic.IsNumeric(CStr(text))
        End Function

        ''' <summary>
        ''' Appends two text inputs and returns the result as another text.  This operation is particularly useful when dealing with unknown text in variables which could accidentally be treated as numbers and get added, instead of getting appended.
        ''' </summary>
        ''' <param name="text1">First part of the text to be appended.</param>
        ''' <param name="text2">Second part of the text to be appended. You can also send an array to append all its items.</param>
        ''' <returns>
        ''' The appended text containing both the specified parts.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function Append(text1 As Primitive, text2 As Primitive) As Primitive
            If text1.IsArray Then
                Dim map1 = text1._arrayMap
                Dim maxKey = map1.Count
                For Each index In map1.Keys
                    Dim id = index.TryGetAsDecimal()
                    If id > maxKey Then maxKey = id
                Next
                map1(maxKey + 1) = text2
                Return text1

            ElseIf text2.IsArray Then
                Dim map = text2._arrayMap
                Return String.Concat(text1, String.Concat(map.Values))
            Else
                Return String.Concat(text1, text2)
            End If
        End Function

        ''' <summary>
        ''' Gets the length of the given text.
        ''' </summary>
        ''' <param name="text">
        ''' The text whose length is needed.
        ''' </param>
        ''' <returns>
        ''' The length of the given text.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Function GetLength(text As Primitive) As Primitive
            If text.IsArray Then Return text._arrayMap.Count
            Return text.AsString().Length
        End Function

        ''' <summary>
        ''' Gets whether or not a given subText is a subset of the larger text.
        ''' </summary>
        ''' <param name="text">
        ''' The larger text within which the sub-text will be searched.
        ''' </param>
        ''' <param name="subText">
        ''' The sub-text to search for.
        ''' </param>
        ''' <returns>
        ''' True if the subtext was found within the given text.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.Boolean)>
        Public Shared Function IsSubText(text As Primitive, subText As Primitive) As Primitive
            If text.IsEmpty OrElse subText.IsEmpty Then Return False
            Return text.AsString().Contains(subText.AsString())
        End Function

        ''' <summary>
        ''' Gets whether or not a given text ends with the specified subText.
        ''' </summary>
        ''' <param name="text">
        ''' The larger text to search within.
        ''' </param>
        ''' <param name="subText">
        ''' The sub-text to search for.
        ''' </param>
        ''' <returns>
        ''' True if the subtext was found at the end of the given text.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.Boolean)>
        Public Shared Function EndsWith(text As Primitive, subText As Primitive) As Primitive
            If text.IsEmpty OrElse subText.IsEmpty Then Return False
            Return text.AsString().EndsWith(subText.AsString())
        End Function

        ''' <summary>
        ''' Gets whether or not a given text starts with the specified subText.
        ''' </summary>
        ''' <param name="text">
        ''' The larger text to search within.
        ''' </param>
        ''' <param name="subText">
        ''' The sub-text to search for.
        ''' </param>
        ''' <returns>
        ''' True if the subtext was found at the start of the given text.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.Boolean)>
        Public Shared Function StartsWith(text As Primitive, subText As Primitive) As Primitive
            If text.IsEmpty OrElse subText.IsEmpty Then Return False
            Return text.AsString().StartsWith(subText.AsString())
        End Function

        ''' <summary>
        ''' Gets whether or not a given text contains the specified subText.
        ''' </summary>
        ''' <param name="text">
        ''' The larger text to search within.
        ''' </param>
        ''' <param name="subText">
        ''' The sub-text to search for.
        ''' </param>
        ''' <returns>
        ''' True if the subtext was found at any posision in the given text, or False otherwise.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.Boolean)>
        Public Shared Function Contains(text As Primitive, subText As Primitive) As Primitive
            If text.IsEmpty OrElse subText.IsEmpty Then Return False
            Return text.AsString().Contains(subText.AsString())
        End Function

        ''' <summary>
        ''' Gets a sub-text from the given text.
        ''' </summary>
        ''' <param name="text">
        ''' The text to derive the sub-text from.
        ''' </param>
        ''' <param name="start">
        ''' Specifies where to start from.
        ''' </param>
        ''' <param name="length">
        ''' Specifies the length of the sub text.
        ''' </param>
        ''' <returns>
        ''' The requested sub-text
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function GetSubText(
                       text As Primitive,
                       start As Primitive,
                       length As Primitive
                   ) As Primitive
            If text.IsEmpty OrElse start.IsEmpty OrElse length.IsEmpty Then Return ""

            Dim strText = text.AsString()
            Dim textLength = strText.Length

            Dim intStart = System.Math.Max(CInt(start), 1)
            intStart = System.Math.Min(intStart, textLength)

            Dim intLength = System.Math.Min(CInt(length), textLength)
            intLength = System.Math.Max(0, intLength)

            If intStart + intLength <= textLength Then
                Return strText.Substring(intStart - 1, intLength)
            End If

            Return strText.Substring(intStart - 1)
        End Function

        ''' <summary>
        ''' Gets a sub-text from the given text from a specified position to the end.
        ''' </summary>
        ''' <param name="text">
        ''' The text to derive the sub-text from.
        ''' </param>
        ''' <param name="start">
        ''' Specifies where to start from.
        ''' </param>
        ''' <returns>
        ''' The requested sub-text.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function GetSubTextToEnd(text As Primitive, start As Primitive) As Primitive
            If text.IsEmpty OrElse start.IsEmpty Then Return ""

            Dim strText = text.AsString()
            Dim intStart = System.Math.Max(CInt(start), 1)
            intStart = System.Math.Min(intStart, strText.Length)

            Return strText.Substring(intStart - 1)
        End Function

        ''' <summary>
        ''' Finds the position where a sub-text appears in the specified text.
        ''' </summary>
        ''' <param name="text">the text to search in.</param>
        ''' <param name="subText">the text to search for.</param>
        ''' <param name="start">the text position to start seacting from</param>
        ''' <param name="isBackward">True if you want to search from start back to the the first position in the text (1), or False if you want to go forward to the end of the text.</param>
        ''' <returns>
        ''' The position at which the sub-text appears in the specified text.  If the text doesn't appear, it returns 0.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Function GetIndexOf(
                         text As Primitive,
                         subText As Primitive,
                         start As Primitive,
                         isBackward As Primitive
                   ) As Primitive

            If text.IsEmpty OrElse subText.IsEmpty Then Return 0
            If Not start.IsNumber OrElse start.AsDecimal < 1 Then Return 0
            If start.AsDecimal > text.AsString.Length + 1 Then Return 0

            If isBackward Then
                Return text.AsString().LastIndexOf(subText, start - 1) + 1
            Else
                Return text.AsString().IndexOf(subText, start - 1) + 1
            End If
        End Function

        ''' <summary>
        ''' Converts the given text to lower case.
        ''' </summary>
        ''' <param name="text">
        ''' The text to convert to lower case.
        ''' </param>
        ''' <returns>
        ''' The lower case version of the given text.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function ConvertToLowerCase(text As Primitive) As Primitive
            If text.IsEmpty Then Return text
            Return CStr(text).ToLower(CultureInfo.CurrentUICulture)
        End Function

        ''' <summary>
        ''' Converts the given text to upper case.
        ''' </summary>
        ''' <param name="text">
        ''' The text to convert to upper case.
        ''' </param>
        ''' <returns>
        ''' The upper case version of the given text.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function ConvertToUpperCase(text As Primitive) As Primitive
            If text.IsEmpty Then Return text
            Return CStr(text).ToUpper(CultureInfo.InvariantCulture)
        End Function

        ''' <summary>
        ''' Given the Unicode character code, gets the corresponding character, which can then be used with regular text.
        ''' </summary>
        ''' <param name="characterCode">
        ''' The character code (Unicode based) for the required character.
        ''' </param>
        ''' <returns>
        ''' A Unicode character that corresponds to the code specified.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function GetCharacter(characterCode As Primitive) As Primitive
            Dim num As Integer = characterCode
            Return ChrW(num).ToString()
        End Function

        ''' <summary>
        ''' Given a Unicode character, gets the corresponding character code.
        ''' </summary>
        ''' <param name="character">
        ''' The character whose code is requested.
        ''' </param>
        ''' <returns>
        ''' A Unicode based code that corresponds to the character specified.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Function GetCharacterCode(character As Primitive) As Primitive
            If character.IsEmpty Then Return 0
            Return AscW(character)
        End Function

        ''' <summary>
        ''' returns the char existing in the given posision in the text
        ''' </summary>
        ''' <param name="text">The input text</param>
        ''' <param name="pos">The posision of the char</param>
        ''' <returns>The char exusted in the given position</returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function GetCharacterAt(text As Primitive, pos As Primitive) As Primitive
            Dim s = CStr(text)
            If s = "" OrElse pos < 1 OrElse pos > s.Length Then Return ""
            Return s(pos - 1)
        End Function

        ''' <summary>
        ''' Returns the new line characters so you can add them to the text to continue writing in a new line
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared ReadOnly Property NewLine As Primitive
            Get
                Return Environment.NewLine
            End Get
        End Property


        ''' <summary>
        ''' changes the char existing in the given posision to the givin value
        ''' </summary>
        ''' <param name="text">the input text</param>
        ''' <param name="pos">The posision of the char</param>
        ''' <returns>a new text with the char changed to the given value. The input text will not change</returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function SetCharacterAt(
                   text As Primitive,
                   pos As Primitive,
                   value As Primitive
            ) As Primitive

            Primitive.SetArrayValue(value, text, pos)
            Return text
        End Function

        ''' <summary>
        ''' Converts the input text to a number
        ''' </summary>
        ''' <param name="text">the input text</param>
        ''' <returns>If text is numeric, returns the numeric value.
        ''' If text contains only one character, returns the ascii code of this character
        ''' Otherwise, returns 0.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Function ToNumber(text As Primitive) As Primitive
            If text.IsNumber Then Return text.AsDecimal
            Dim x = text.AsString()
            If x.Length > 1 Then Return New Primitive(0D)
            Return AscW(x)
        End Function

        ''' <summary>
        ''' Converts all characters of the input text to lower case.
        ''' </summary>
        ''' <param name="text">the input text</param>
        ''' <returns>a lower-case text</returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function ToLower(text As Primitive) As Primitive
            Return text.AsString().ToLower()
        End Function

        ''' <summary>
        ''' Converts all characters of the input text to upper case.
        ''' </summary>
        ''' <param name="text">the input text</param>
        ''' <returns>an upper-case text</returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function ToUpper(text As Primitive) As Primitive
            Return text.AsString().ToUpper()
        End Function

        ''' <summary>
        '''Removes all leading and trailing white-space characters from the given text
        '''Wite-space chars iclude spaces, tabs, and line symbols.
        ''' </summary>
        ''' <param name="text">the input text</param>
        ''' <returns>the trimmed string</returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function Trim(text As Primitive) As Primitive
            If text.IsEmpty Then Return text
            If text.IsArray Then Return text
            Return text.AsString().Trim()
        End Function

        ''' <summary>
        '''Converts the given value or array to a string.
        ''' </summary>
        ''' <param name="value">the input value</param>
        ''' <returns>The string representaion of the value. For example, the array string can be of the form {1, 2, 3}</returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function ToStr(value As Primitive) As Primitive
            If value.IsArray Then
                If value.IsEmpty Then Return "{}"

                Dim sb As New System.Text.StringBuilder("{")
                Dim map = value._arrayMap
                Dim n = map.Count - 1
                Dim keys = map.Keys
                Dim values = map.Values
                Dim isNormalArray = True

                For i = 0 To n
                    If keys(i) <> i + 1 Then
                        isNormalArray = False
                        Exit For
                    End If
                Next

                For i = 0 To n
                    If isNormalArray Then
                        sb.Append(ToStr(values(i)).AsString())
                    Else
                        sb.AppendFormat("[{0}]={1}", keys(i), ToStr(values(i)))
                    End If
                    If i < n Then sb.Append(", ")
                Next

                sb.Append("}")
                Return sb.ToString()
            Else
                Return value.AsString()
            End If
        End Function

        ''' <summary>
        '''Returns true if the given text is empty, or returns false otherwise.
        ''' </summary>
        ''' <param name="text">the input text</param>
        ''' <returns>True or False</returns>
        <WinForms.ReturnValueType(VariableType.Boolean)>
        Public Shared Function IsEmpty(text As Primitive) As Primitive
            Return text.IsEmpty
        End Function

        ''' <summary>
        ''' Splits the given text at the given separator.
        ''' </summary>
        ''' <param name="text">The input text</param>
        ''' <param name="separator">One character or more to split the text at. The separator will not appear in the result. You can also send an array to use its items as separators.</param>
        ''' <param name="trim">Use True to trim white spaces from the start and end of the separated strings</param>
        ''' <param name="removeEmpty">Use True to remove empty strings from the result</param>
        ''' <returns>An array containing the splitted items</returns>
        <WinForms.ReturnValueType(VariableType.Array)>
        Public Shared Function Split(
                        text As Primitive,
                        separator As Primitive,
                        trim As Primitive,
                        removeEmpty As Primitive
                    ) As Primitive
            If text.IsEmpty Then Return ""

            Dim arr, separators As String()
            If separator.IsArray Then
                Dim items = separator._arrayMap.Values
                Dim n = items.Count - 1
                separators = New String(n) {}

                For i = 0 To n
                    separators(i) = items(i).AsString()
                Next

            Else
                separators = {separator.AsString()}
            End If

            Dim options = If(removeEmpty, StringSplitOptions.RemoveEmptyEntries, StringSplitOptions.None)
            arr = text.AsString().Split(separators, options)

            If CBool(trim) Then
                For i = 0 To arr.Count - 1
                    arr(i) = arr(i).Trim()
                Next
            End If

            Return New Primitive(arr, removeEmpty)
        End Function
    End Class
End Namespace
