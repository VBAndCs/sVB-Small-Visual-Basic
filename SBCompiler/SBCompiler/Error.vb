Namespace Microsoft.SmallBasic
    Public Class [Error]

        Public ReadOnly Property Line As Integer
        Public ReadOnly Property subLine As Integer
        Public ReadOnly Property Column As Integer
        Public ReadOnly Property Description As String

        Public Sub New(token As Token, description As String)
            Me.New(token.Line, token.subLine, token.Column, description)
        End Sub

        Public Sub New(line As Integer, subLine As Integer, column As Integer, description As String)
            _Line = line
            _subLine = subLine
            _Column = column
            _Description = description
        End Sub

        Public Overrides Function ToString() As String
            Return $"{Line},{Column}: {Description}"
        End Function
    End Class
End Namespace
