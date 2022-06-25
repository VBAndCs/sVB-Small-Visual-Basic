Imports System.Windows.Automation.Peers

Namespace Microsoft.Nautilus.Text.Editor.Automation.Implementation
    Friend Class TextEditorAutomationPeer
        Inherits FrameworkElementAutomationPeer

        Private _textView As AvalonTextView

        Public Sub New(textView As AvalonTextView)
            MyBase.New(textView)
            _textView = textView
        End Sub

        Protected Overrides Function GetClassNameCore() As String
            Return "AvalonTextView"
        End Function

        Protected Overrides Function GetAutomationControlTypeCore() As AutomationControlType
            Return AutomationControlType.Text
        End Function

        Protected Overrides Function GetNameCore() As String
            Return "Nautilus Text View"
        End Function

        Protected Overrides Function GetAutomationIdCore() As String
            Return "nautilusTextView"
        End Function

        Public Overrides Function GetPattern(patternInterface As PatternInterface) As Object
            Select Case patternInterface
                Case PatternInterface.Text
                    Return New TextEditorPatternProvider(_textView)

                Case PatternInterface.Value
                    Return _textView

                Case Else
                    Return MyBase.GetPattern(patternInterface)

            End Select
        End Function
    End Class
End Namespace
