Imports System
Imports System.Globalization

Namespace Microsoft.SmallBasic
    <Serializable>
    Public Structure TokenInfo
        Public Property Line As Integer

        Public Property Column As Integer

        Public ReadOnly Property EndColumn As Integer
            Get
                If Text = "" Then Return Column
                Return Column + Text.Length
            End Get
        End Property

        Public ReadOnly Property NormalizedText As String
            Get
                If Text = "" Then Return ""
                Return Text.ToLower(CultureInfo.CurrentUICulture)
            End Get
        End Property

        Public Property Text As String

        Public Property Token As Token

        Public Property TokenType As TokenType

        Public Shared ReadOnly Property Illegal As New TokenInfo With {
            .Line = 0,
            .Column = 0,
            .Token = Token.Illegal,
            .TokenType = TokenType.Illegal
        }

        Public Overrides Function ToString() As String
            Return $"{Line},{Column}: {Text}, {Token}"
        End Function
    End Structure
End Namespace
