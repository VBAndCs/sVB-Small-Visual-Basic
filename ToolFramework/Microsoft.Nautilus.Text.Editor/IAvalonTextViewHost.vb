Imports System.Windows.Controls

Namespace Microsoft.Nautilus.Text.Editor
    Public Interface IAvalonTextViewHost
        Property IsHorizontalScrollBarVisible As Boolean

        Property IsVerticalScrollBarVisible As Boolean

        Property IsReadOnly As Boolean

        ReadOnly Property TextView As IAvalonTextView

        ReadOnly Property HostControl As Control

        Property ScaleFactor As Double

        Property WordWrapStyle As WordWrapStyles

        Function GetTextViewMargin(marginName As String) As ITextViewMargin
    End Interface
End Namespace
