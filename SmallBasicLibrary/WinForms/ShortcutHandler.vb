Imports Microsoft.SmallVisualBasic.Library
Imports KeyNames = System.Windows.Input.Key

Namespace WinForms
    Public Class ShortcutHandler
        Public ReadOnly Handler As SmallVisualBasicCallback
        Public ReadOnly Ctrl, Shift, Alt As Boolean
        Public ReadOnly key As KeyNames

        Public Sub New(shortcut As String, handler As SmallVisualBasicCallback)
            Me.Handler = handler
            Ctrl = shortcut.Contains("Ctrl+")
            Shift = shortcut.Contains("Shift+")
            Alt = shortcut.Contains("Alt+")
            Dim k = shortcut.Substring(shortcut.LastIndexOf("+") + 1)
            Select Case k
                Case "-"
                    key = KeyNames.OemMinus
                Case "+"
                    key = KeyNames.OemPlus
                Case "["
                    key = KeyNames.OemOpenBrackets
                Case "]"
                    key = KeyNames.OemCloseBrackets
                Case "0"
                    key = KeyNames.D0
                Case "1"
                    key = KeyNames.D1
                Case "2"
                    key = KeyNames.D2
                Case "3"
                    key = KeyNames.D3
                Case "4"
                    key = KeyNames.D4
                Case "5"
                    key = KeyNames.D5
                Case "6"
                    key = KeyNames.D6
                Case "7"
                    key = KeyNames.D7
                Case "8"
                    key = KeyNames.D8
                Case "9"
                    key = KeyNames.D9
                Case "A"
                    key = KeyNames.A
                Case "B"
                    key = KeyNames.B
                Case "Back"
                    key = KeyNames.Back
                Case "C"
                    key = KeyNames.C
                Case "D"
                    key = KeyNames.D
                Case "Delete"
                    key = KeyNames.Delete
                Case "Down"
                    key = KeyNames.Down
                Case "E"
                    key = KeyNames.E
                Case "Esc"
                    key = KeyNames.Escape
                Case "F"
                    key = KeyNames.F
                Case "F1"
                    key = KeyNames.F1
                Case "F2"
                    key = KeyNames.F2
                Case "F3"
                    key = KeyNames.F3
                Case "F4"
                    key = KeyNames.F4
                Case "F5"
                    key = KeyNames.F5
                Case "F6"
                    key = KeyNames.F6
                Case "F7"
                    key = KeyNames.F7
                Case "F8"
                    key = KeyNames.F8
                Case "F9"
                    key = KeyNames.F9
                Case "F10"
                    key = KeyNames.F10
                Case "F11"
                    key = KeyNames.F11
                Case "F12"
                    key = KeyNames.F12
                Case "G"
                    key = KeyNames.G
                Case "H"
                    key = KeyNames.H
                Case "I"
                    key = KeyNames.I
                Case "J"
                    key = KeyNames.J
                Case "K"
                    key = KeyNames.K
                Case "L"
                    key = KeyNames.L
                Case "Left"
                    key = KeyNames.Left
                Case "M"
                    key = KeyNames.M
                Case "N"
                    key = KeyNames.N
                Case "O"
                    key = KeyNames.O
                Case "P"
                    key = KeyNames.P
                Case "Q"
                    key = KeyNames.Q
                Case "R"
                    key = KeyNames.R
                Case "Right"
                    key = KeyNames.Right
                Case "S"
                    key = KeyNames.S
                Case "T"
                    key = KeyNames.T
                Case "Tab"
                    key = KeyNames.Tab
                Case "U"
                    key = KeyNames.U
                Case "Up"
                    key = KeyNames.Up
                Case "V"
                    key = KeyNames.V
                Case "W"
                    key = KeyNames.W
                Case "X"
                    key = KeyNames.X
                Case "Y"
                    key = KeyNames.Y
                Case "Z"
                    key = KeyNames.Z
            End Select
        End Sub
    End Class

End Namespace

