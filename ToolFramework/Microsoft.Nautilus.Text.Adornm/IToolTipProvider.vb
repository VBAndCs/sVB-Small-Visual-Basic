Namespace Microsoft.Nautilus.Text.Adornments
    Public Interface IToolTipProvider
        Sub ShowToolTip(span As ITextSpan, toolTipText As String, style As ToolTipStyles)

        Sub ShowToolTip(span As ITextSpan, toolTipText As String)

        Sub ClearToolTip()
    End Interface
End Namespace
