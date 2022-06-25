Imports System.Windows.Automation.Peers

Namespace Microsoft.Nautilus.Text.Editor.Automation.Implementation
    Friend Class TextViewHostAutomationPeer
        Inherits FrameworkElementAutomationPeer

        Private Const HostWindowClassName As String = "Microsoft.Nautilus.Text.Editor.AvalonTextViewHost"
        Private automationID As String

        Public Sub New(host As AvalonTextViewHost)
            MyBase.New(host)

            Dim peer = UIElementAutomationPeer.CreatePeerForElement(CType(host.TextView, AvalonTextView))
            automationID = peer.GetAutomationId()
        End Sub

        Protected Overrides Function GetClassNameCore() As String
            Return "Microsoft.Nautilus.Text.Editor.AvalonTextViewHost"
        End Function

        Protected Overrides Function GetAutomationIdCore() As String
            Return automationID
        End Function

        Protected Overrides Function IsContentElementCore() As Boolean
            Return False
        End Function

        Protected Overrides Function IsControlElementCore() As Boolean
            Return False
        End Function

    End Class
End Namespace
