Namespace Microsoft.SmallBasic
    Public Class [Error]

        Public ReadOnly Property Line As Integer

        Public ReadOnly Property Column As Integer

        Public ReadOnly Property Description As String

        Public Sub New(tokenInfo As TokenInfo, description As String)
            Me.New(tokenInfo.Line, tokenInfo.Column, description)
        End Sub

        Public Sub New(line As Integer, column As Integer, description As String)
            _Line = line
            _Column = column
            _Description = description
        End Sub

        Public Overrides Function ToString() As String
            Return $"{Line},{Column}: {Description}"
        End Function
    End Class
End Namespace
