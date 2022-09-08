Imports System.Globalization
Imports Microsoft.SmallBasic.Library

Namespace WinForms

    <SmallBasicType>
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
        ''' Gets whether or not a given text starts with the specified subText.
        ''' </summary>
        ''' <param name="subText">
        ''' The sub-text to search for.
        ''' </param>
        ''' <returns>
        ''' True if the subtext was found at any posision in the given text.
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
        ''' <param name="subText">
        ''' The text to search for.
        ''' </param>
        ''' <returns>
        ''' The position at which the sub-text appears in the specified text.  If the text doesn't appear, it returns 0.
        ''' </returns>
        <ReturnValueType(VariableType.Double)>
        <ExMethod>
        Public Shared Function IndexOf(text As Primitive, subText As Primitive) As Primitive
            Return Library.Text.GetIndexOf(text, subText)
        End Function

        ''' <summary>
        ''' Converts the given text to lower case.
        ''' </summary>
        ''' <returns>
        ''' The lower case version of the given text.
        ''' </returns>
        <ReturnValueType(VariableType.String)>
        <ExProperty>
        Public Shared Function LowerCase(text As Primitive) As Primitive
            If text.IsEmpty Then Return text
            Return CStr(text).ToLower(CultureInfo.CurrentUICulture)
        End Function

        ''' <summary>
        ''' Converts the given text to upper case.
        ''' </summary>
        ''' <returns>
        ''' The upper case version of the given text.
        ''' </returns>
        <ReturnValueType(VariableType.String)>
        <ExProperty>
        Public Shared Function UpperCase(text As Primitive) As Primitive
            If text.IsEmpty Then Return text
            Return CStr(text).ToUpper(CultureInfo.InvariantCulture)
        End Function

        ''' <summary>
        ''' returns the char existing in the given posision in the text
        ''' </summary>
        ''' <param name="pos">The posision of the char</param>
        ''' <returns>The char exusted in the given position</returns>
        <ReturnValueType(VariableType.String)>
        <ExMethod>
        Public Shared Function GetCharacterAt(text As Primitive, pos As Primitive) As Primitive
            Return Library.Text.GetCharacterAt(text, pos)
        End Function

        ''' <summary>
        ''' changes the char existing in the given posision to the givin value
        ''' </summary>
        ''' <param name="pos">The posision of the char</param>
        ''' <returns>a new text with the char changed to the given value. The current text will not change</returns>
        <ReturnValueType(VariableType.String)>
        <ExMethod>
        Public Shared Function SetCharacterAt(text As Primitive, pos As Primitive, value As Primitive) As Primitive
            Return Library.Text.SetCharacterAt(text, pos, value)
        End Function
    End Class
End Namespace
