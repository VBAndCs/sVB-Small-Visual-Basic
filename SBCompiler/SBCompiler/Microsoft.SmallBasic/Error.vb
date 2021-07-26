Namespace Microsoft.SmallBasic
    Public Class [Error]

        Public ReadOnly Property Line As Integer

        Public ReadOnly Property Column As Integer

        Public ReadOnly Property Description As String

        Public Sub New(ByVal tokenInfo As TokenInfo, ByVal description As String)
            Me.New(tokenInfo.Line, tokenInfo.Column, description)
        End Sub

        Public Sub New(ByVal line As Integer, ByVal column As Integer, ByVal description As String)
            _Line = line
            _Column = column
            _description = description
        End Sub

        Public Overrides Function ToString() As String
            Return $"{Line},{Column}: {Description}"
        End Function
    End Class
End Namespace
