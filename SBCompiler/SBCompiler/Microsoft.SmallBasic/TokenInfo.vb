Imports System
Imports System.Globalization

Namespace Microsoft.SmallBasic
    <Serializable>
    Public Structure TokenInfo
        Public Shared ChangeTextFunc As Func(Of Token, String, String)

        Private Shared _illegal As New TokenInfo With {
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

        Dim _text As String

        Public Property Text As String
            Get
                Return _text
            End Get

            Set(value As String)
                If ChangeTextFunc Is Nothing Then
                    _text = value
                Else
                    _text = ChangeTextFunc(_Token, value)
                End If
            End Set
        End Property

        Private _Token As Token

        Public Property Token As Token
            Get
                Return _Token
            End Get

            Set
                _Token = Value
                If ChangeTextFunc IsNot Nothing Then _text = ChangeTextFunc(_Token, _text)
            End Set
        End Property

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
