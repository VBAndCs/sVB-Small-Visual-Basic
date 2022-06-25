Imports System.Collections.Generic

Namespace SuperClassifier
    Public Class Identifiers
        Private _keywords As IList(Of String) = New List(Of String)

        Public Property PrefixCharacters As String

        Public Property BodyCharacters As String

        Public Property IsCaseSensitive As Boolean

        Public ReadOnly Property Keywords As IList(Of String)
            Get
                Return _keywords
            End Get
        End Property
    End Class
End Namespace
