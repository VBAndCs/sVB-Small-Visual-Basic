Namespace Microsoft.Nautilus.Text
    Public Interface IReadOnlyRegionHandle
        ReadOnly Property EdgeInsertionMode As EdgeInsertionMode

        ReadOnly Property Span As ITextSpan

        Sub Remove()
    End Interface
End Namespace
