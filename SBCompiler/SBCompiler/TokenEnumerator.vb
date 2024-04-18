Imports System.Collections.Generic

Namespace Microsoft.SmallVisualBasic
    Public Class TokenEnumerator
        Friend TokensList As List(Of Token)
        Private _currentIndex As Integer

        Public Property LineNumber As Integer

        Public ReadOnly Property IsEnd As Boolean
            Get
                Return _currentIndex = TokensList.Count
            End Get
        End Property


        Public ReadOnly Property IsEndOrComment As Boolean
            Get
                If _currentIndex = TokensList.Count Then Return True
                Return _Current.ParseType = ParseType.Comment
            End Get
        End Property

        Public ReadOnly Property Current As Token

        Public Sub New(tokenList As List(Of Token), Optional startAt As Integer = 0)
            TokensList = tokenList

            If tokenList.Count > 0 Then
                _currentIndex = startAt
                _Current = tokenList(startAt)
            End If
        End Sub

        Public Function PeekNext() As Token
            If _currentIndex < TokensList.Count - 1 Then
                Return TokensList(_currentIndex + 1)
            End If

            Return Token.Illegal
        End Function

        Public Function MoveNext() As Boolean
            If _currentIndex >= TokensList.Count Then
                Return False
            End If

            _currentIndex += 1

            If _currentIndex < TokensList.Count Then
                _Current = TokensList(_currentIndex)
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
                _Current = TokensList(_currentIndex)
                Return True
            End If

            Return False
        End Function

        Public Sub Reset()
            _currentIndex = 0
            _Current = If(_currentIndex < TokensList.Count, TokensList(_currentIndex), Nothing)
        End Sub

    End Class


End Namespace
