Imports System
Imports System.Globalization

Namespace Evaluator
    <Serializable>
    Public Structure Token
        Public Line As Integer
        Public Column As Integer
        Public Text As String
        Public Type As TokenType

        Public ReadOnly Property EndColumn As Integer
            Get
                If Text = "" Then Return Column
                Return Column + Text.Length
            End Get
        End Property

        Public ReadOnly Property LCaseText As String
            Get
                If Text = "" Then Return ""
                Return Text.ToLower(CultureInfo.CurrentUICulture)
            End Get
        End Property

        Public Shared ReadOnly Illegal As New Token With {
            .Line = 0,
            .Column = 0,
            .Type = TokenType.Illegal
        }

        Public ReadOnly Property IsIllegal As Boolean
            Get
                Return Type = TokenType.Illegal
            End Get
        End Property

        Public Overrides Function ToString() As String
            Return $"{Line},{Column}: '{Text}', {Type}"
        End Function
    End Structure


    Public Class TokenEnumerator
        Private _tokenList As List(Of Token)
        Private _currentIndex As Integer

        Public Property LineNumber As Integer

        Public ReadOnly Property IsEnd As Boolean
            Get
                Return _currentIndex = _tokenList.Count
            End Get
        End Property

        Public ReadOnly Property Current As Token

        Public Sub New(tokenList As List(Of Token))
            _tokenList = tokenList

            If tokenList.Count > 0 Then
                _Current = tokenList(0)
            End If
        End Sub

        Public Function PeekNext() As Token
            If _currentIndex < _tokenList.Count - 1 Then
                Return _tokenList(_currentIndex + 1)
            End If

            Return Token.Illegal
        End Function

        Public Function MoveNext() As Boolean
            If _currentIndex >= _tokenList.Count Then
                Return False
            End If

            _currentIndex += 1

            If _currentIndex < _tokenList.Count Then
                _Current = _tokenList(_currentIndex)
                Return True
            End If

            Return False
        End Function

        Public Function StepBack() As Boolean
            If _currentIndex < 0 Then
                Return False
            End If

            _currentIndex -= 1

            If _currentIndex >= 0 Then
                _Current = _tokenList(_currentIndex)
                Return True
            End If

            Return False
        End Function

        Public Sub Reset()
            _currentIndex = 0
            _Current = If(_currentIndex < _tokenList.Count, _tokenList(_currentIndex), Nothing)
        End Sub

    End Class

End Namespace
