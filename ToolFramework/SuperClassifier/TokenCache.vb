Imports System.Collections.Generic

Namespace SuperClassifier
    Friend Class TokenCache
        Friend ReadOnly startIndex As Integer
        Friend ReadOnly length As Integer
        Friend ReadOnly tokensCoveringCache As List(Of Token)

        Friend Sub New(startIndex As Integer, length As Integer, tokensCoveringCache As List(Of Token))
            Me.startIndex = startIndex
            Me.length = length
            Me.tokensCoveringCache = tokensCoveringCache
        End Sub
    End Class
End Namespace
