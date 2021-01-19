Imports System.Globalization
Imports System.Reflection.Emit
Imports Microsoft.SmallBasic.Completion

Namespace Microsoft.SmallBasic.Statements
    Public Class GotoStatement
        Inherits Statement

        Public GotoToken As TokenInfo
        Public Label As TokenInfo

        Public Overrides Sub EmitIL(ByVal scope As CodeGenScope)
            Dim label = scope.Labels(Me.Label.NormalizedText)
            scope.ILGenerator.Emit(OpCodes.Br, label)
        End Sub

        Public Overrides Sub PopulateCompletionItems(ByVal completionBag As CompletionBag, ByVal line As Integer, ByVal column As Integer, ByVal globalScope As Boolean)
            If Label.Token = Token.Illegal OrElse column < Label.Column Then
                CompletionHelper.FillAllGlobalItems(completionBag, globalScope)
            Else
                CompletionHelper.FillLabels(completionBag)
            End If
        End Sub

        Public Overrides Function ToString() As String
            Return String.Format(CultureInfo.CurrentUICulture, "{0} {1}" & VisualBasic.Constants.vbCrLf, New Object(1) {GotoToken.Text, Label.Text})
        End Function
    End Class
End Namespace
