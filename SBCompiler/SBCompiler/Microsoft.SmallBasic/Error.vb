Namespace Microsoft.SmallBasic
    Public Class [Error]
        Private _line As Integer
        Private _column As Integer
        Private _description As String

        Public ReadOnly Property Line As Integer
            Get
                Return _line
            End Get
        End Property

        Public ReadOnly Property Column As Integer
            Get
                Return _column
            End Get
        End Property

        Public ReadOnly Property Description As String
            Get
                Return _description
            End Get
        End Property

        Public Sub New(ByVal tokenInfo As TokenInfo, ByVal description As String)
            Me.New(tokenInfo.Line, tokenInfo.Column, description)
        End Sub

        Public Sub New(ByVal line As Integer, ByVal column As Integer, ByVal description As String)
            _line = line
            _column = column
            _description = description
        End Sub

        Public Overrides Function ToString() As String
            Return $"{Line},{Column}: {Description}"
        End Function
    End Class
End Namespace
