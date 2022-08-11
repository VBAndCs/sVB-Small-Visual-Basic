Imports Microsoft.SmallBasic.Library

<SmallBasicType>
Public NotInheritable Class Chars

    ''' <summary>
    '''     Represents a carriage return character.
    ''' </summary>
    Public Shared ReadOnly Property Cr As Primitive = ChrW(13)

    ''' <summary>
    '''     Represents a line feed character.
    ''' </summary>
    Public Shared ReadOnly Property Lf As Primitive = ChrW(10)

    ''' <summary>
    '''     Represents a line feed character.
    ''' </summary>
    Public Shared ReadOnly Property CrLf As Primitive = Chr(13) + ChrW(10)

    ''' <summary>
    '''     Represents a backspace character.
    ''' </summary>
    Public Shared ReadOnly Property Back As Primitive = ChrW(8)

    ''' <summary>
    '''     Represents a form feed character for print functions.
    ''' </summary>
    Public Shared ReadOnly Property FormFeed As Primitive = ChrW(12)

    ''' <summary>
    '''     Represents a tab character.
    ''' </summary>
    Public Shared ReadOnly Property Tab As Primitive = ChrW(9)

    ''' <summary>
    '''     Represents a vertical tab character.
    ''' </summary>
    Public Shared ReadOnly Property VerticalTab As Primitive = ChrW(11)

    ''' <summary>
    '''     Represents a null character.
    ''' </summary>
    Public Shared ReadOnly Property Null As Primitive = ChrW(0)

    ''' <summary>
    '''     Represents a double-quote character.
    ''' </summary>
    Public Shared ReadOnly Property Quote As Primitive = """"c
End Class

