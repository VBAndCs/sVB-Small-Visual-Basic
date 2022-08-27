
Namespace Microsoft.Nautilus.Text.Editor
    Public Interface ITextView
        Inherits IPropertyOwner
        ReadOnly Property ViewScroller As IViewScroller

        ReadOnly Property FormattedTextLines As IFormattedTextLineCollection

        ReadOnly Property FirstVisibleCharacterPosition As Integer

        Property WordWrapStyle As WordWrapStyles

        ReadOnly Property Caret As ITextCaret

        ReadOnly Property Selection As ITextSelection

        ReadOnly Property TextBuffer As ITextBuffer

        ReadOnly Property TextSnapshot As ITextSnapshot

        ReadOnly Property MaxTextWidth As Double

        Property ViewportLeft As Double

        ReadOnly Property ViewportTop As Double

        ReadOnly Property ViewportRight As Double

        ReadOnly Property ViewportBottom As Double

        ReadOnly Property ViewportWidth As Double

        ReadOnly Property ViewportHeight As Double

        Event LayoutChanged As EventHandler(Of TextViewLayoutChangedEventArgs)
        Event ViewportLeftChanged As EventHandler
        Event ViewportHeightChanged As EventHandler
        Event ViewportWidthChanged As EventHandler
        Event WordWrapStyleChanged As EventHandler
        Event MouseHover As EventHandler(Of MouseHoverEventArgs)

        Sub DisplayTextLineContainingCharacter(position As Integer, verticalDistance As Double, relativeTo As ViewRelativePosition)

        Sub DisplayAllContent()

        Function GetTextElementSpan(position As Integer) As Span
    End Interface
End Namespace
