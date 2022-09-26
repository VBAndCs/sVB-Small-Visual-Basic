Imports System
Imports System.Globalization

Namespace Microsoft.SmallVisualBasic
    <Serializable>
    Public Structure Token
        Public Line As Integer
        Public subLine As Integer
        Public Column As Integer
        Public Text As String
        Public Parent As Statements.Statement
        Public Type As TokenType
        Public ParseType As ParseType
        Public SymbolType As Completion.CompletionItemType
        Public Comment As String

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
            .Type = TokenType.Illegal,
            .ParseType = ParseType.Illegal
        }

        Public ReadOnly Property IsIllegal As Boolean
            Get
                Return Type = TokenType.Illegal
            End Get
        End Property

        Public ReadOnly Property SubroutineName As String
            Get
                Dim subroutine = Statements.SubroutineStatement.GetSubroutine(Me)
                If subroutine Is Nothing Then Return ""
                Return subroutine.Name.Text
            End Get
        End Property

        Public Function Contains(
                        line As Integer,
                        column As Integer,
                        Optional encludeEndColumn As Boolean = False
                 ) As Boolean

            If line < 0 Then Return True ' to allow any token to be addee to the completion bag

            Return Me.Line = line AndAlso
                        column >= Me.Column AndAlso (
                            column < Me.EndColumn OrElse
                            (encludeEndColumn AndAlso column = EndColumn)
                        )
        End Function

        Public Shared Operator =(token1 As Token, token2 As Token) As Boolean
            Return token1.Line = token2.Line AndAlso token1.Column = token2.Column
        End Operator

        Public Shared Operator <>(token1 As Token, token2 As Token) As Boolean
            Return token1.Line <> token2.Line OrElse token1.Column <> token2.Column
        End Operator

        Public Function IsBefore(line As Integer, column As Integer) As Boolean
            If Me.Type = TokenType.Illegal Then Return False
            If Me.Line > line Then Return False
            If Me.Line < line Then Return True
            Return Me.EndColumn < column
        End Function

        Public Function IsAfter(line As Integer, column As Integer) As Boolean
            If Me.Type = TokenType.Illegal Then Return True
            If Me.Line < line Then Return True
            If Me.Line > line Then Return False
            Return Me.Column > column
        End Function

        Public Overrides Function ToString() As String
            Return $"{Line},{Column}: '{Text}', {Type}"
        End Function
    End Structure
End Namespace
