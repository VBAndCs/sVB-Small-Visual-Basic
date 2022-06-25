Imports System.IO

Namespace Microsoft.Nautilus.Text.Editor
    Public Interface ITextEditorFactory
        Property TrackEditors As Boolean

        Function CreateTextView() As IAvalonTextView

        Function CreateTextView(textBuffer As ITextBuffer) As IAvalonTextView

        Function CreateTextViewHost() As IAvalonTextViewHost

        Function CreateTextViewHost(textBuffer As ITextBuffer) As IAvalonTextViewHost

        Function CreateTextViewHost(avalonTextView As IAvalonTextView) As IAvalonTextViewHost

        Sub TagEditor(editor As IAvalonTextView, tag As String)

        Function ReportLiveEditors(writer As TextWriter) As Integer
    End Interface
End Namespace
