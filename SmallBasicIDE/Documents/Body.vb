Imports System.Windows.Documents

Namespace Microsoft.SmallBasic
    Public Class Body
        Inherits Span

        Public Sub New(text As String)
            FontSize = 13.0
            Inlines.Add(New Run(text))
        End Sub
    End Class
End Namespace
