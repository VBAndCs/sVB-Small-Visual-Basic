Imports System.Windows
Imports System.Windows.Documents
Imports System.Windows.Media

Namespace Microsoft.SmallVisualBasic
    Public Class Heading
        Inherits Span

        Public Sub New(text As String)
            FontSize = 24.0
            FontWeight = FontWeights.Bold
            Foreground = New SolidColorBrush(Color.FromRgb(52, 109, 132))
            Inlines.Add(New Run(text))
            Inlines.Add(New LineBreak())
        End Sub
    End Class
End Namespace
