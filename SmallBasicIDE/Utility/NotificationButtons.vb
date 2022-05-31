Imports System

Namespace Microsoft.SmallBasic.Utility
    <Flags>
    Public Enum NotificationButtons
        Close = &H1
        Cancel = &H2
        Retry = &H4
        No = &H8
        Yes = &H10
        OK = &H20
        Custom = &H8000
    End Enum
End Namespace
