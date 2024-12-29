Imports System.Globalization
Imports System.Reflection.Emit
Imports Microsoft.SmallVisualBasic.Completion
Imports Microsoft.SmallVisualBasic.Engine

Namespace Microsoft.SmallVisualBasic.Statements
    Public Class GotoStatement
        Inherits Statement

        Public GotoToken As Token
        Public Label As Token
        Public subroutine As SubroutineStatement

        Public Sub New()
        End Sub

        ''' <summary>
        ''' Creates a Goto statement that jumps to a line number.
        ''' It is used for debugging only, but not allowed in code.
        ''' </summary>
        ''' <param name="line"></param>
        Public Sub New(line As Integer)
            Label = Token.Illegal
            Label.Line = line
        End Sub

        Public Overrides Function GetStatementAt(lineNumber As Integer) As Statement
            If lineNumber < StartToken.Line Then Return Nothing
            If lineNumber <= Label.Line Then Return Me
            Return Nothing
        End Function

        Public Overrides Sub AddSymbols(symbolTable As SymbolTable)
            MyBase.AddSymbols(symbolTable)
            GotoToken.Parent = Me
            Label.Parent = Me
            Label.SymbolType = CompletionItemType.Label
            If Not Label.IsIllegal Then symbolTable.AddIdentifier(Label)
        End Sub

        Public Overrides Sub EmitIL(scope As CodeGenScope)
            If scope.ForGlobalHelp OrElse Me.Label.IsIllegal Then Return

            Dim label = scope.Labels(Me.Label.LCaseText)
            scope.ILGenerator.Emit(OpCodes.Br, label)
        End Sub

        Public Overrides Sub PopulateCompletionItems(completionBag As CompletionBag, line As Integer, column As Integer, globalScope As Boolean)
            completionBag.CompletionItems.Clear()
            CompletionHelper.FillLabels(completionBag)
        End Sub

        Public Overrides Function ToString() As String
            Return $"{GotoToken.Text} {Label.Text}" & vbCrLf
        End Function

        Public Overrides Function Execute(runner As ProgramRunner) As Statement
            Return Me
        End Function

        Public Overrides Function ToVB() As String
            Return $"{GotoToken.Text} {Label.Text}" & vbCrLf
        End Function
    End Class
End Namespace
