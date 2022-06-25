Imports System.Collections.Generic
Imports System.Windows
Imports System.Windows.Media

Namespace Microsoft.Nautilus.Text.Classification.DataExports
    Public NotInheritable Class ClassificationFormat

        Public ReadOnly Property ClassificationTypes As New List(Of String)

        Public Property ForegroundBrush As Brush

        Public Property BackgroundBrush As Brush

        Public Property Culture As Integer?

        Public Property FontHintingSize As Double?

        Public Property FontRenderingSize As Double?

        Public ReadOnly Property TextEffects As New TextEffectCollection

        Public ReadOnly Property TextDecorations As New TextDecorationCollection

        Public Property FontFamily As String

        Public Property FontStyle As FontStyle

        Public Property FontWeight As FontWeight

        Public Property FontStretch As FontStretch

        Public Property FallbackFontFamily As String

        Public Property After As String

        Public Property Before As String

    End Class
End Namespace
