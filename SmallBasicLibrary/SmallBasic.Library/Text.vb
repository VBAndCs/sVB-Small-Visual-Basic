Imports System.Globalization

Namespace Library
    ''' <summary>
    ''' The Text object provides helpful operations for working with Text.
    ''' </summary>
    <SmallBasicType>
    Public NotInheritable Class Text
        Public Shared Function IsNumeric(value As Primitive) As Primitive
            Return VisualBasic.IsNumeric(CStr(value))
        End Function

        ''' <summary>
        ''' Appends two text inputs and returns the result as another text.  This operation is particularly useful when dealing with unknown text in variables which could accidentally be treated as numbers and get added, instead of getting appended.
        ''' </summary>
        ''' <param name="text1">
        ''' First part of the text to be appended.
        ''' </param>
        ''' <param name="text2">
        ''' Second part of the text to be appended.
        ''' </param>
        ''' <returns>
        ''' The appended text containing both the specified parts.
        ''' </returns>
        Public Shared Function Append(text1 As Primitive, text2 As Primitive) As Primitive
            Return String.Concat(text1, text2)
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
        Public Shared Function GetLength(text As Primitive) As Primitive
            If text.IsEmpty Then Return 0
            Return CStr(text).Length
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
        Public Shared Function IsSubText(text As Primitive, subText As Primitive) As Primitive
            If text.IsEmpty OrElse subText.IsEmpty Then Return False
            Return CStr(text).Contains(subText)
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
        Public Shared Function EndsWith(text As Primitive, subText As Primitive) As Primitive
            If text.IsEmpty OrElse subText.IsEmpty Then Return False
            Return CStr(text).EndsWith(subText)
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
        Public Shared Function StartsWith(text As Primitive, subText As Primitive) As Primitive
            If text.IsEmpty OrElse subText.IsEmpty Then Return False
            Return CStr(text).StartsWith(subText)
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
        Public Shared Function GetSubText(text As Primitive, start As Primitive, length As Primitive) As Primitive
            If text.IsEmpty OrElse start.IsEmpty OrElse length.IsEmpty Then Return ""

            Dim strText = CStr(text)
            Dim intStart = CInt(start)
            Dim intLength = CInt(length)

            If intStart > strText.Length OrElse intStart < 1 Then
                Return ""
            End If

            If intStart + intLength <= strText.Length Then
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
        Public Shared Function GetSubTextToEnd(text As Primitive, start As Primitive) As Primitive
            If text.IsEmpty OrElse start.IsEmpty Then Return ""

            Dim strText = CStr(text)
            Dim intStart = CInt(start)

            If intStart > strText.Length OrElse intStart < 1 Then
                Return ""
            End If

            Return strText.Substring(intStart - 1)
        End Function

        ''' <summary>
        ''' Finds the position where a sub-text appears in the specified text.
        ''' </summary>
        ''' <param name="text">
        ''' The text to search in.
        ''' </param>
        ''' <param name="subText">
        ''' The text to search for.
        ''' </param>
        ''' <returns>
        ''' The position at which the sub-text appears in the specified text.  If the text doesn't appear, it returns 0.
        ''' </returns>
        Public Shared Function GetIndexOf(text As Primitive, subText As Primitive) As Primitive
            If text.IsEmpty OrElse subText.IsEmpty Then Return 0
            Return CStr(text).IndexOf(subText) + 1
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
        Public Shared Function GetCharacter(characterCode As Primitive) As Primitive
            Dim num As Integer = characterCode
            Return ChrW(num).ToString()
        End Function

        Public Shared Function GetCharacterAt(text As Primitive, pos As Primitive) As Primitive
            Dim s = CStr(text)
            If s = "" OrElse pos < 1 OrElse pos > s.Length Then Return ""
            Return s(pos - 1)
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
        Public Shared Function GetCharacterCode(character As Primitive) As Primitive
            If character.IsEmpty Then Return 0
            Return AscW(character)
        End Function
    End Class
End Namespace
