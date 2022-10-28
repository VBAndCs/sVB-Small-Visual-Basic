Imports Microsoft.SmallVisualBasic.Library

Namespace WinForms

    ''' <summary>
    ''' Contains spme useful non-printed characters, such as the Escap and new line characters. 
    ''' </summary>
    <SmallBasicType>
    Public NotInheritable Class Chars

        ''' <summary>
        '''     Represents a carriage return character.
        ''' </summary>
        <ReturnValueType(VariableType.String)>
        Public Shared ReadOnly Property Cr As Primitive = ChrW(13)

        ''' <summary>
        '''     Represents a line feed character.
        ''' </summary>
        <ReturnValueType(VariableType.String)>
        Public Shared ReadOnly Property Lf As Primitive = ChrW(10)

        ''' <summary>
        '''     Represents a line feed character.
        ''' </summary>
        <ReturnValueType(VariableType.String)>
        Public Shared ReadOnly Property CrLf As Primitive = Chr(13) + ChrW(10)

        ''' <summary>
        '''     Represents a backspace character.
        ''' </summary>
        <ReturnValueType(VariableType.String)>
        Public Shared ReadOnly Property Back As Primitive = ChrW(8)

        ''' <summary>
        '''     Represents a form feed character for print functions.
        ''' </summary>
        <ReturnValueType(VariableType.String)>
        Public Shared ReadOnly Property FormFeed As Primitive = ChrW(12)

        ''' <summary>
        '''     Represents a tab character.
        ''' </summary>
        <ReturnValueType(VariableType.String)>
        Public Shared ReadOnly Property Tab As Primitive = ChrW(9)

        ''' <summary>
        '''     Represents a vertical tab character.
        ''' </summary>
        <ReturnValueType(VariableType.String)>
        Public Shared ReadOnly Property VerticalTab As Primitive = ChrW(11)

        ''' <summary>
        '''     Represents a null character.
        ''' </summary>
        <ReturnValueType(VariableType.String)>
        Public Shared ReadOnly Property Null As Primitive = ChrW(0)

        ''' <summary>
        '''     Represents a double-quote character.
        ''' </summary>
        <ReturnValueType(VariableType.String)>
        Public Shared ReadOnly Property Quote As Primitive = """"c

        ''' <summary>
        ''' Given the Unicode character code, gets the corresponding character, which can then be used with regular text.
        ''' </summary>
        ''' <param name="characterCode">
        ''' The character code (Unicode based) for the required character.
        ''' </param>
        ''' <returns>
        ''' A Unicode character that corresponds to the code specified.
        ''' </returns>
        <ReturnValueType(VariableType.String)>
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
        <ReturnValueType(VariableType.Double)>
        Public Shared Function GetCharacterCode(character As Primitive) As Primitive
            If character.IsEmpty Then Return 0
            Return AscW(character)
        End Function

        ''' <summary>
        ''' Checks if the given character is a letter in any human language.
        ''' </summary>
        ''' <param name="character">the character to ckeck.</param>
        ''' <returns>True of False.</returns>
        <ReturnValueType(VariableType.Boolean)>
        Public Shared Function IsLetter(character As Primitive) As Primitive
            Return Char.IsLetter(character)
        End Function

        ''' <summary>
        ''' Checks if the given character is a digit (0-9).
        ''' </summary>
        ''' <param name="character">the character to ckeck.</param>
        ''' <returns>True of False.</returns>
        <ReturnValueType(VariableType.Boolean)>
        Public Shared Function IsDigit(character As Primitive) As Primitive
            Return Char.IsDigit(character)
        End Function

    End Class


End Namespace
