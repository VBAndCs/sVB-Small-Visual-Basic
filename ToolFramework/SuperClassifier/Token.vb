Imports Microsoft.Nautilus.Text

Namespace SuperClassifier
    Friend Class Token
        Friend TokenStart As Integer
        Friend ReadOnly TokenLength As Integer
        Friend ReadOnly TokenString As String
        Friend ReadOnly Classification As String

        Friend ReadOnly Property TokenEnd As Integer
            Get
                Return TokenStart + TokenLength
            End Get
        End Property

        Friend ReadOnly Property Span As Span
            Get
                Return New Span(TokenStart, TokenLength)
            End Get
        End Property

        Friend Sub New(start As Integer, length As Integer, text As String, classification As String)
            Me.TokenStart = start
            Me.TokenLength = length
            Me.TokenString = text
            Me.Classification = classification
        End Sub

        Friend Sub ChangeTokenStart(threshold As Integer, delta As Integer)
            If TokenStart >= threshold Then TokenStart += delta
        End Sub

        Public Overrides Function ToString() As String
            Return $"({TokenStart},{TokenLength}) {Classification} ""{TokenString}"""
        End Function
    End Class
End Namespace
