Imports System.Windows
Imports System.Windows.Automation
Imports System.Windows.Automation.Provider

Namespace Microsoft.Nautilus.Text.Editor.Automation.Implementation
    Friend Class TextEditorPatternProvider
        Implements ITextProvider

        ReadOnly Property ITextProvider_SupportedTextSelection As SupportedTextSelection Implements ITextProvider.SupportedTextSelection
            Get
                Return SupportedTextSelection.Single
            End Get
        End Property

        ReadOnly Property ITextProvider_DocumentRange As ITextRangeProvider Implements ITextProvider.DocumentRange
            Get
                Return New TextRangeProvider(_textView, New Span(0, _textView.TextSnapshot.Length))
            End Get
        End Property

        Friend ReadOnly Property TextView As IAvalonTextView

        Public Sub New(textView As IAvalonTextView)
            _TextView = textView
        End Sub

        Private Function GetSelection() As ITextRangeProvider() Implements ITextProvider.GetSelection
            If Not _textView.Selection.IsEmpty Then
                Return New TextRangeProvider(0) {New TextRangeProvider(_textView, _textView.Selection.ActiveSnapshotSpan)}
            End If
            Return Nothing
        End Function

        Private Function GetVisibleRanges() As ITextRangeProvider() Implements ITextProvider.GetVisibleRanges
            Dim formattedTextLines1 As IFormattedTextLineCollection = _textView.FormattedTextLines
            Return New ITextRangeProvider(0) {
                New TextRangeProvider(_textView, formattedTextLines1.FirstFullyVisibleLine.LineSpan)
            }
        End Function

        Private Function RangeFromChild(childElement As IRawElementProviderSimple) As ITextRangeProvider Implements ITextProvider.RangeFromChild
            Throw New NotImplementedException("We don't support child elements in Text yet.")
        End Function

        Private Function RangeFromPoint(screenLocation As Point) As ITextRangeProvider Implements ITextProvider.RangeFromPoint
            Throw New NotImplementedException
        End Function
    End Class
End Namespace
