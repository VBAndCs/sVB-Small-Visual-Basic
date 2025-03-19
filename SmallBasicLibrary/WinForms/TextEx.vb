Imports System.Globalization
Imports Microsoft.SmallVisualBasic.Library

Namespace WinForms

    <SmallVisualBasicType>
    <HideFromIntellisense>
    Public NotInheritable Class TextEx

        ''' <summary>
        ''' Formats the string by replacing [1], [2], ... [n] by items from the values array.
        ''' </summary>
        ''' <param name="values">An array its elements will be used to replace [1], [2],... [n] strings if found in the text</param>
        ''' <returns>The formated string after substituting [1], [2],... [n] with elements from the values array</returns>
        <ReturnValueType(VariableType.String)>
        <ExMethod>
        Public Shared Function Format(text As Primitive, values As Primitive) As Primitive
            Return Library.Text.Format(text, values)
        End Function


        ''' <summary>
        ''' Checks if the string contains a numeric value
        ''' </summary>
        ''' <returns>True if text is a number, False otherwise</returns>
        <ReturnValueType(VariableType.Boolean)>
        <ExProperty>
        Public Shared Function GetIsNumeric(text As Primitive) As Primitive
            Return Library.Text.IsNumeric(text)
        End Function

        ''' <summary>
        ''' Appends two text inputs and returns the result as another text.  This operation is particularly useful when dealing with unknown text in variables which could accidentally be treated as numbers and get added, instead of getting appended.
        ''' </summary>
        ''' <param name="text">
        ''' Second part of the text to be appended.
        ''' </param>
        ''' <returns>
        ''' The appended text containing both the specified parts.
        ''' </returns>
        <ReturnValueType(VariableType.String)>
        <ExMethod>
        Public Shared Function Append(text1 As Primitive, text As Primitive) As Primitive
            Return Library.Text.Append(text1, text)
        End Function

        ''' <summary>
        ''' Gets the length of the given text.
        ''' </summary>
        ''' <returns>
        ''' The length of the given text.
        ''' </returns>
        <ReturnValueType(VariableType.Double)>
        <WinForms.ExProperty>
        Public Shared Function GetLength(text As Primitive) As Primitive
            Return Library.Text.GetLength(text)
        End Function

        ''' <summary>
        ''' Gets whether or not a given text ends with the specified subText.
        ''' </summary>
        ''' <param name="subText">
        ''' The sub-text to search for.
        ''' </param>
        ''' <returns>
        ''' True if the subtext was found at the end of the given text.
        ''' </returns>
        <ReturnValueType(VariableType.Boolean)>
        <ExMethod>
        Public Shared Function EndsWith(text As Primitive, subText As Primitive) As Primitive
            Return Library.Text.EndsWith(text, subText)
        End Function

        ''' <summary>
        ''' Gets whether or not a given text starts with the specified subText.
        ''' </summary>
        ''' <param name="subText">
        ''' The sub-text to search for.
        ''' </param>
        ''' <returns>
        ''' True if the subtext was found at the start of the given text.
        ''' </returns>
        <ReturnValueType(VariableType.Boolean)>
        <ExMethod>
        Public Shared Function StartsWith(text As Primitive, subText As Primitive) As Primitive
            Return Library.Text.StartsWith(text, subText)
        End Function

        ''' <summary>
        ''' Gets whether or not a given text contains the specified subText.
        ''' </summary>
        ''' <param name="subText">
        ''' The sub-text to search for.
        ''' </param>
        ''' <returns>
        ''' True if the subtext was found at any posision in the given text, or False otherwise.
        ''' </returns>
        <ReturnValueType(VariableType.Boolean)>
        <ExMethod>
        Public Shared Function Contains(text As Primitive, subText As Primitive) As Primitive
            Return Library.Text.Contains(text, subText)
        End Function

        ''' <summary>
        ''' Gets a sub-text from the given text.
        ''' </summary>
        ''' <param name="start">
        ''' Specifies where to start from.
        ''' </param>
        ''' <param name="length">
        ''' Specifies the length of the sub text.
        ''' </param>
        ''' <returns>
        ''' The requested sub-text
        ''' </returns>
        <ReturnValueType(VariableType.String)>
        <ExMethod>
        Public Shared Function SubText(text As Primitive, start As Primitive, length As Primitive) As Primitive
            Return Library.Text.GetSubText(text, start, length)
        End Function

        ''' <summary>
        ''' Gets a sub-text from the given text from a specified position to the end.
        ''' </summary>
        ''' <param name="start">
        ''' Specifies where to start from.
        ''' </param>
        ''' <returns>
        ''' The requested sub-text.
        ''' </returns>
        <ReturnValueType(VariableType.String)>
        <ExMethod>
        Public Shared Function SubTextToEnd(text As Primitive, start As Primitive) As Primitive
            Return Library.Text.GetSubTextToEnd(text, start)
        End Function

        ''' <summary>
        ''' Finds the position where a sub-text appears in the specified text.
        ''' </summary>
        ''' <param name="subText">the text to search for.</param>
        ''' <param name="start">the text position to start seacting from</param>
        ''' <param name="isBackward">True if you want to search from start back to the the first position in the text (1), or False if you want to go forward to the end of the text.</param>
        ''' <returns>
        ''' The position at which the sub-text appears in the specified text.  If the text doesn't appear, it returns 0.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.Double)>
        <ExMethod>
        <HideFromIntellisense>
        Public Shared Function GetIndexOf(
                         text As Primitive,
                         subText As Primitive,
                         start As Primitive,
                         isBackward As Primitive
                   ) As Primitive

            Return Library.Text.IndexOf(text, subText, start, isBackward)
        End Function

        ''' <summary>
        ''' Finds the position where a sub-text appears in the specified text.
        ''' </summary>
        ''' <param name="subText">the text to search for.</param>
        ''' <param name="start">the text position to start seacting from</param>
        ''' <param name="isBackward">True if you want to search from start back to the the first position in the text (1), or False if you want to go forward to the end of the text.</param>
        ''' <returns>
        ''' The position at which the sub-text appears in the specified text.  If the text doesn't appear, it returns 0.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.Double)>
        <ExMethod>
        Public Shared Function IndexOf(
                         text As Primitive,
                         subText As Primitive,
                         start As Primitive,
                         isBackward As Primitive
                   ) As Primitive

            Return Library.Text.IndexOf(text, subText, start, isBackward)
        End Function

        ''' <summary>
        ''' Converts the given text to lower case.
        ''' </summary>
        ''' <returns>
        ''' The lower case version of the given text.
        ''' </returns>
        <ReturnValueType(VariableType.String)>
        <ExProperty>
        Public Shared Function GetLowerCase(text As Primitive) As Primitive
            Return Library.Text.ToLower(text)
        End Function


        ''' <summary>
        ''' Converts the given text to lower case.
        ''' </summary>
        ''' <returns>
        ''' The lower case version of the given text.
        ''' </returns>
        <ReturnValueType(VariableType.String)>
        <ExMethod>
        Public Shared Function ToLower(text As Primitive) As Primitive
            Return Library.Text.ToLower(text)
        End Function


        ''' <summary>
        ''' Converts the given text to upper case.
        ''' </summary>
        ''' <returns>
        ''' The upper case version of the given text.
        ''' </returns>
        <ReturnValueType(VariableType.String)>
        <ExProperty>
        Public Shared Function GetUpperCase(text As Primitive) As Primitive
            Return Library.Text.ToUpper(text)
        End Function


        ''' <summary>
        ''' Converts the given text to upper case.
        ''' </summary>
        ''' <returns>
        ''' The upper case version of the given text.
        ''' </returns>
        <ReturnValueType(VariableType.String)>
        <ExMethod>
        Public Shared Function ToUpper(text As Primitive) As Primitive
            Return Library.Text.ToUpper(text)
        End Function

        ''' <summary>
        ''' Gets the character existing in the given posision in the text
        ''' </summary>
        ''' <param name="pos">The posision of the character</param>
        ''' <returns>The character existed in the given position</returns>
        <ReturnValueType(VariableType.String)>
        <ExMethod>
        Public Shared Function GetCharacterAt(text As Primitive, pos As Primitive) As Primitive
            Return Library.Text.GetCharacterAt(text, pos)
        End Function

        ''' <summary>
        ''' Changes the char existing in the given posision to the givin value
        ''' </summary>
        ''' <param name="pos">The posision of the char</param>
        ''' <returns>a new text with the char changed to the given value. The current text will not change</returns>
        <ReturnValueType(VariableType.String)>
        <ExMethod>
        Public Shared Function SetCharacterAt(text As Primitive, pos As Primitive, value As Primitive) As Primitive
            Return Library.Text.SetCharacterAt(text, pos, value)
        End Function

        ''' <summary>
        '''Removes all leading and trailing white-space characters from the given text
        '''Wite-space chars iclude spaces, tabs, and line symbols.
        ''' </summary>
        ''' <returns>the trimmed string</returns>
        <ExMethod>
        <ReturnValueType(VariableType.String)>
        Public Shared Function Trim(text As Primitive) As Primitive
            Return Library.Text.Trim(text)
        End Function

        ''' <summary>
        '''Returns true if the current text is empty, or returns false otherwise.
        ''' </summary>
        ''' <returns>True or False</returns>
        <ExProperty>
        <WinForms.ReturnValueType(VariableType.Boolean)>
        Public Shared Function GetIsEmpty(text As Primitive) As Primitive
            Return text.IsEmpty
        End Function

        ''' <summary>
        ''' Splits the current text at the given separator.
        ''' </summary>
        ''' <param name="separator">One character or more to split the text at. The separator will not appear in the result.</param>
        ''' <param name="trim">Use True to trim white spaces from the start and end of the separated strings</param>
        ''' <param name="removeEmpty">Use True to remove empty strings from the result</param>
        ''' <returns>An array containing the splitted items</returns>
        <WinForms.ReturnValueType(VariableType.Array)>
        <ExMethod>
        Public Shared Function Split(text As Primitive, separator As Primitive, trim As Primitive, removeEmpty As Primitive) As Primitive
            Return Library.Text.Split(text, separator, trim, removeEmpty)
        End Function

        ''' <summary>
        ''' Checks if the current text is a letter in any language.
        ''' </summary>
        ''' <returns>
        ''' True in the text consists of one character and it is a letter in any language, itherwise False.
        ''' Note that this property is meant to b used with a single character. It the text length is > 1, this method will always return false. In such case, you should use a for loop to get each charcter in the text and check it individually, if thi is what you need.
        ''' </returns>
        <ReturnValueType(VariableType.Boolean)>
        <ExProperty>
        Public Shared Function GetIsLetter(text As Primitive) As Primitive
            Return Chars.IsLetter(text)
        End Function

        ''' <summary>
        ''' Checks if the given character is a digit (0-9).
        ''' </summary>
        ''' <returns>True or False.</returns>
        <ReturnValueType(VariableType.Boolean)>
        <ExProperty>
        Public Shared Function GetIsDigit(text As Primitive) As Primitive
            Return Chars.IsDigit(text)
        End Function

        ''' <summary>
        ''' Converts the current hexadecimal string to a decimal number. 
        ''' You can use the Math.Hex method to convert a decimal number to a hexadecimal number.
        ''' </summary>
        ''' <returns>The decimal value the hexa number if it is valid, or an empty string otherwise.</returns>
        <WinForms.ReturnValueType(VariableType.Double)>
        <ExMethod>
        Public Shared Function ToDecimal(hex As Primitive) As Primitive
            Return Convert.ToInt32(hex, 16)
        End Function


        ''' <summary>
        ''' Adds spaces on the left of the current text so that its length equals the given total width. This is useful when you want to right-align texts.
        ''' Note that if the length of the current text is greater than or egual to the given length, it will be displayed as it is and will not be trimmed.
        ''' </summary>
        ''' <param name="totalWidth">The desired total length of the text.</param>
        ''' <returns>a new text that has the given total length (at least).</returns>
        <WinForms.ReturnValueType(VariableType.String)>
        <ExMethod>
        Public Shared Function PadLeft(
                   text As Primitive,
                   totalWidth As Primitive
            ) As Primitive

            Return New Primitive(text.AsString().PadLeft(totalWidth))
        End Function

        ''' <summary>
        ''' Adds spaces on the right of the current text so that its length equals the given total width. This is useful when you want to left-align texts.
        ''' Note that if the length of the current text is greater than or egual to the given length, it will be displayed as it is and will not be trimmed.
        ''' </summary>
        ''' <param name="totalWidth">The desired total length of the text.</param>
        ''' <returns>a new text that has the given total length (at least).</returns>
        <WinForms.ReturnValueType(VariableType.String)>
        <ExMethod>
        Public Shared Function PadRight(
                   text As Primitive,
                   totalWidth As Primitive
            ) As Primitive

            Return New Primitive(text.AsString().PadRight(totalWidth))
        End Function

        ''' <summary>
        ''' Converts the current text to a number
        ''' </summary>
        ''' <returns>If text is numeric, returns the numeric value.
        ''' If text contains only one character, returns the ascii code of this character
        ''' Otherwise, returns 0.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.Double)>
        <ExMethod>
        Public Shared Function ToNumber(text As Primitive) As Primitive
            Return Library.Text.ToNumber(text)
        End Function


        ''' <summary>
        ''' Creates a new date from the current text if it has a valid date format for the given culture.
        ''' </summary>
        ''' <param name="cultureName">
        ''' The culture name used to format the date, like "en-US" for English (United States) culture, "ar-EG" for Arabic (Egypt) culture, and "ar-SA" for Arabic (Saudi Arabia) culture.
        ''' ''' Send an empty string "" to use the local culture of the user's system.
        ''' </param>
        ''' <returns>a new date or empty string</returns>
        <ReturnValueType(VariableType.Date)>
        <ExMethod>
        Public Shared Function ToDate(dateText As Primitive, cultureName As Primitive) As Primitive
            Return [Date].FromCulture(dateText, cultureName)
        End Function

        ''' <summary>
        ''' Converts the input text to a duration if it has a valid format.
        ''' </summary>
        ''' <returns>If text is a valid duration, returns the the duration value.
        ''' Otherwise, returns a 0 duration.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.Date)>
        <ExMethod>
        Public Shared Function ToDuration(text As Primitive) As Primitive
            Return Library.Text.ToDuration(text)
        End Function



        ''' <summary>
        '''Converts the current text to an array, if it has a valid array format.
        ''' </summary>
        ''' <returns>an array that is constructed from the text if it has a valid array format, otherwise an empty array.</returns>
        <WinForms.ReturnValueType(VariableType.Array)>
        <ExMethod>
        Public Shared Function ToArray(value As Primitive) As Primitive
            Return Text.ToArray(value)
        End Function


        ''' <summary>
        ''' Repeats the current text for the given number of times.
        ''' For examle, when you repeat "aB" 3 times, it returns "aBaBaB".
        ''' </summary>
        ''' <param name="count">the number of times to repeat the text.</param>
        ''' <returns>the repeated text</returns>
        <WinForms.ReturnValueType(VariableType.String)>
        <ExMethod>
        Public Shared Function Repeat(text As Primitive, count As Primitive) As Primitive
            Return Library.Text.Repeat(text, count)
        End Function

    End Class
End Namespace
