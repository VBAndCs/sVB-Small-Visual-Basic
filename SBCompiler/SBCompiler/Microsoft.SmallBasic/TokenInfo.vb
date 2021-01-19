Imports System
Imports System.Globalization

Namespace Microsoft.SmallBasic
    <Serializable>
    Public Structure TokenInfo
        Private Shared _illegal As TokenInfo = New TokenInfo With {
            .Line = 0,
            .Column = 0,
            .Token = Token.Illegal,
            .TokenType = TokenType.Illegal
        }
        Public Property Line As Integer
        Public Property Column As Integer

        Public ReadOnly Property EndColumn As Integer
            Get

                If Not Equals(Text, Nothing) Then
                    Return Column + Text.Length
                End If

                Return Column
            End Get
        End Property

        Public ReadOnly Property NormalizedText As String
            Get

                If Not Equals(Text, Nothing) Then
                    Return Text.ToLower(CultureInfo.CurrentUICulture)
                End If

                Return Nothing
            End Get
        End Property

        Public Property Text As String
        Public Property Token As Token
        Public Property TokenType As TokenType

        Public Shared ReadOnly Property Illegal As TokenInfo
            Get
                Return _illegal
            End Get
        End Property

        Public Overrides Function ToString() As String
            Return $"{Line},{Column}: {Text}, {Token}"
        End Function
    End Structure
End Namespace
