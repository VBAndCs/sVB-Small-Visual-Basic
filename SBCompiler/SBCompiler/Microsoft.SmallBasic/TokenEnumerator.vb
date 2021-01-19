Imports System.Collections.Generic

Namespace Microsoft.SmallBasic
    Public Class TokenEnumerator
        Private _tokenList As List(Of TokenInfo)
        Private _currentIndex As Integer
        Private _current As TokenInfo
        Public Property LineNumber As Integer

        Public ReadOnly Property IsEndOfList As Boolean
            Get
                Return _currentIndex = _tokenList.Count
            End Get
        End Property

        Public ReadOnly Property IsEndOfNonCommentList As Boolean
            Get

                If _currentIndex <> _tokenList.Count Then
                    Return _current.TokenType = TokenType.Comment
                End If

                Return True
            End Get
        End Property

        Public ReadOnly Property Current As TokenInfo
            Get
                Return _current
            End Get
        End Property

        Public Sub New(ByVal tokenList As List(Of TokenInfo))
            _tokenList = tokenList

            If tokenList.Count > 0 Then
                _current = tokenList(0)
            End If
        End Sub

        Public Function PeekNext() As TokenInfo
            If _currentIndex < _tokenList.Count - 1 Then
                Return _tokenList(_currentIndex + 1)
            End If

            Return TokenInfo.Illegal
        End Function

        Public Function MoveNext() As Boolean
            If _currentIndex >= _tokenList.Count Then
                Return False
            End If

            _currentIndex += 1

            If _currentIndex < _tokenList.Count Then
                _current = _tokenList(_currentIndex)
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
                _current = _tokenList(_currentIndex)
                Return True
            End If

            Return False
        End Function

        Public Sub Reset()
            _currentIndex = 0
            _current = If(_currentIndex < _tokenList.Count, _tokenList(_currentIndex), Nothing)
        End Sub
    End Class
End Namespace
