Namespace Microsoft.Nautilus.Text.Operations
    Public Interface ITextStructureNavigatorFactory
        Function CreateTextStructureNavigator(textBuffer As ITextBuffer) As ITextStructureNavigator

        Function CreateTextStructureNavigator(textBuffer As ITextBuffer, contentType As String) As ITextStructureNavigator
    End Interface
End Namespace
