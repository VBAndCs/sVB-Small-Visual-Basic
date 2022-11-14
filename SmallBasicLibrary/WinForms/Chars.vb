Imports Microsoft.SmallVisualBasic.Library

Namespace WinForms

    ''' <summary>
    ''' Contains some useful non-printed characters, such as the Escape and new line characters. 
    ''' </summary>
    <SmallVisualBasicType>
    Public NotInheritable Class Chars

        ''' <summary>
        ''' Represents a carriage return character, which is the end of line.
        ''' </summary>
        <ReturnValueType(VariableType.String)>
        Public Shared ReadOnly Property Cr As Primitive = ChrW(13)

        ''' <summary>
        ''' Represents a line feed character, which is the start of a new line. In some cases, this character a lone can be used to represent the new line.
        ''' </summary>
        <ReturnValueType(VariableType.String)>
        Public Shared ReadOnly Property Lf As Primitive = ChrW(10)

        ''' <summary>
        ''' Represents a carriage return and a line feed character. Together, they indicate the presence of a new line.
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
            Dim code As Integer = characterCode
            Return Char.ConvertFromUtf32(code)
        End Function

        ''' <summary>
        ''' Given a Unicode character, gets the corresponding character code.
        ''' </summary>
        ''' <param name="character">The character whose code is requested.</param>
        ''' <returns>
        ''' A Unicode based code that corresponds to the character specified.
        ''' </returns>
        <ReturnValueType(VariableType.Double)>
        Public Shared Function GetCharacterCode(character As Primitive) As Primitive
            If character.IsEmpty Then Return 0
            Return Char.ConvertToUtf32(character.AsString(), 0)
        End Function

        ''' <summary>
        ''' Checks if the given character is a letter in any language.
        ''' </summary>
        ''' <param name="character">the character to check.</param>
        ''' <returns>True or False.</returns>
        <ReturnValueType(VariableType.Boolean)>
        Public Shared Function IsLetter(character As Primitive) As Primitive
            Dim c = character.AsString()
            If c.Length <> 1 Then Return False
            Return Char.IsLetter(c)
        End Function

        ''' <summary>
        ''' Checks if the given character is a digit (0-9).
        ''' </summary>
        ''' <param name="character">the character to check.</param>
        ''' <returns>True or False.</returns>
        <ReturnValueType(VariableType.Boolean)>
        Public Shared Function IsDigit(character As Primitive) As Primitive
            Dim c = character.AsString()
            If c.Length <> 1 Then Return False
            Return Char.IsDigit(c(0))
        End Function

    End Class


End Namespace
