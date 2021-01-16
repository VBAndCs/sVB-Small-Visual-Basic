Imports Microsoft.SmallBasic.Library

'
' Summary:
'     Specifies the possible key values on a keyboard.
<SmallBasicType>
Public NotInheritable Class Keys

    '
    ' Summary:
    '     No key pressed.
    Public Shared ReadOnly Property None As Primitive = 0
    '
    ' Summary:
    '     The Cancel key.
    Public Shared ReadOnly Property Cancel As Primitive = 1
    '
    ' Summary:
    '     The Backspace key.
    Public Shared ReadOnly Property Back As Primitive = 2
    '
    ' Summary:
    '     The Tab key.
    Public Shared ReadOnly Property Tab As Primitive = 3
    '
    ' Summary:
    '     The Linefeed key.
    Public Shared ReadOnly Property LineFeed As Primitive = 4
    '
    ' Summary:
    '     The Clear key.
    Public Shared ReadOnly Property Clear As Primitive = 5
    '
    ' Summary:
    '     The Return key.
    Public Shared ReadOnly Property [Return] As Primitive = 6
    '
    ' Summary:
    '     The Enter key.
    Public Shared ReadOnly Property Enter As Primitive = 6
    '
    ' Summary:
    '     The Pause key.
    Public Shared ReadOnly Property Pause As Primitive = 7
    '
    ' Summary:
    '     The Caps Lock key.
    Public Shared ReadOnly Property Capital As Primitive = 8
    '
    ' Summary:
    '     The Caps Lock key.
    Public Shared ReadOnly Property CapsLock As Primitive = 8
    '
    ' Summary:
    '     The IME Kana mode key.
    Public Shared ReadOnly Property KanaMode As Primitive = 9
    '
    ' Summary:
    '     The IME Hangul mode key.
    Public Shared ReadOnly Property HangulMode As Primitive = 9
    '
    ' Summary:
    '     The IME Junja mode key.
    Public Shared ReadOnly Property JunjaMode As Primitive = 10
    '
    ' Summary:
    '     The IME Final mode key.
    Public Shared ReadOnly Property FinalMode As Primitive = 11
    '
    ' Summary:
    '     The IME Hanja mode key.
    Public Shared ReadOnly Property HanjaMode As Primitive = 12
    '
    ' Summary:
    '     The IME Kanji mode key.
    Public Shared ReadOnly Property KanjiMode As Primitive = 12
    '
    ' Summary:
    '     The ESC key.
    Public Shared ReadOnly Property Escape As Primitive = 13
    '
    ' Summary:
    '     The IME Convert key.
    Public Shared ReadOnly Property ImeConvert As Primitive = 14
    '
    ' Summary:
    '     The IME NonConvert key.
    Public Shared ReadOnly Property ImeNonConvert As Primitive = 15
    '
    ' Summary:
    '     The IME Accept key.
    Public Shared ReadOnly Property ImeAccept As Primitive = 16
    '
    ' Summary:
    '     The IME Mode change request.
    Public Shared ReadOnly Property ImeModeChange As Primitive = 17
    '
    ' Summary:
    '     The Spacebar key.
    Public Shared ReadOnly Property Space As Primitive = 18
    '
    ' Summary:
    '     The Page Up key.
    Public Shared ReadOnly Property Prior As Primitive = 19
    '
    ' Summary:
    '     The Page Up key.
    Public Shared ReadOnly Property PageUp As Primitive = 19
    '
    ' Summary:
    '     The Page Down key.
    Public Shared ReadOnly Property [Next] As Primitive = 20
    '
    ' Summary:
    '     The Page Down key.
    Public Shared ReadOnly Property PageDown As Primitive = 20
    '
    ' Summary:
    '     The End key.
    Public Shared ReadOnly Property [End] As Primitive = 21
    '
    ' Summary:
    '     The Home key.
    Public Shared ReadOnly Property Home As Primitive = 22
    '
    ' Summary:
    '     The Left Arrow key.
    Public Shared ReadOnly Property Left As Primitive = 23
    '
    ' Summary:
    '     The Up Arrow key.
    Public Shared ReadOnly Property Up As Primitive = 24
    '
    ' Summary:
    '     The Right Arrow key.
    Public Shared ReadOnly Property Right As Primitive = 25
    '
    ' Summary:
    '     The Down Arrow key.
    Public Shared ReadOnly Property Down As Primitive = 26
    '
    ' Summary:
    '     The Select key.
    Public Shared ReadOnly Property [Select] As Primitive = 27
    '
    ' Summary:
    '     The Print key.
    Public Shared ReadOnly Property Print As Primitive = 28
    '
    ' Summary:
    '     The Execute key.
    Public Shared ReadOnly Property Execute As Primitive = 29
    '
    ' Summary:
    '     The Print Screen key.
    Public Shared ReadOnly Property Snapshot As Primitive = 30
    '
    ' Summary:
    '     The Print Screen key.
    Public Shared ReadOnly Property PrintScreen As Primitive = 30
    '
    ' Summary:
    '     The Insert key.
    Public Shared ReadOnly Property Insert As Primitive = 31
    '
    ' Summary:
    '     The Delete key.
    Public Shared ReadOnly Property Delete As Primitive = 32
    '
    ' Summary:
    '     The Help key.
    Public Shared ReadOnly Property Help As Primitive = 33
    '
    ' Summary:
    '     The 0 (zero) key.
    Public Shared ReadOnly Property D0 As Primitive = 34
    '
    ' Summary:
    '     The 1 (one) key.
    Public Shared ReadOnly Property D1 As Primitive = 35
    '
    ' Summary:
    '     The 2 key.
    Public Shared ReadOnly Property D2 As Primitive = 36
    '
    ' Summary:
    '     The 3 key.
    Public Shared ReadOnly Property D3 As Primitive = 37
    '
    ' Summary:
    '     The 4 key.
    Public Shared ReadOnly Property D4 As Primitive = 38
    '
    ' Summary:
    '     The 5 key.
    Public Shared ReadOnly Property D5 As Primitive = 39
    '
    ' Summary:
    '     The 6 key.
    Public Shared ReadOnly Property D6 As Primitive = 40
    '
    ' Summary:
    '     The 7 key.
    Public Shared ReadOnly Property D7 As Primitive = 41
    '
    ' Summary:
    '     The 8 key.
    Public Shared ReadOnly Property D8 As Primitive = 42
    '
    ' Summary:
    '     The 9 key.
    Public Shared ReadOnly Property D9 As Primitive = 43
    '
    ' Summary:
    '     The A key.
    Public Shared ReadOnly Property A As Primitive = 44
    '
    ' Summary:
    '     The B key.
    Public Shared ReadOnly Property B As Primitive = 45
    '
    ' Summary:
    '     The C key.
    Public Shared ReadOnly Property C As Primitive = 46
    '
    ' Summary:
    '     The D key.
    Public Shared ReadOnly Property D As Primitive = 47
    '
    ' Summary:
    '     The E key.
    Public Shared ReadOnly Property E As Primitive = 48
    '
    ' Summary:
    '     The F key.
    Public Shared ReadOnly Property F As Primitive = 49
    '
    ' Summary:
    '     The G key.
    Public Shared ReadOnly Property G As Primitive = 50
    '
    ' Summary:
    '     The H key.
    Public Shared ReadOnly Property H As Primitive = 51
    '
    ' Summary:
    '     The I key.
    Public Shared ReadOnly Property I As Primitive = 52
    '
    ' Summary:
    '     The J key.
    Public Shared ReadOnly Property J As Primitive = 53
    '
    ' Summary:
    '     The K key.
    Public Shared ReadOnly Property K As Primitive = 54
    '
    ' Summary:
    '     The L key.
    Public Shared ReadOnly Property L As Primitive = 55
    '
    ' Summary:
    '     The M key.
    Public Shared ReadOnly Property M As Primitive = 56
    '
    ' Summary:
    '     The N key.
    Public Shared ReadOnly Property N As Primitive = 57
    '
    ' Summary:
    '     The O key.
    Public Shared ReadOnly Property O As Primitive = 58
    '
    ' Summary:
    '     The P key.
    Public Shared ReadOnly Property P As Primitive = 59
    '
    ' Summary:
    '     The Q key.
    Public Shared ReadOnly Property Q As Primitive = 60
    '
    ' Summary:
    '     The R key.
    Public Shared ReadOnly Property R As Primitive = 61
    '
    ' Summary:
    '     The S key.
    Public Shared ReadOnly Property S As Primitive = 62
    '
    ' Summary:
    '     The T key.
    Public Shared ReadOnly Property T As Primitive = 63
    '
    ' Summary:
    '     The U key.
    Public Shared ReadOnly Property U As Primitive = 64
    '
    ' Summary:
    '     The V key.
    Public Shared ReadOnly Property V As Primitive = 65
    '
    ' Summary:
    '     The W key.
    Public Shared ReadOnly Property W As Primitive = 66
    '
    ' Summary:
    '     The X key.
    Public Shared ReadOnly Property X As Primitive = 67
    '
    ' Summary:
    '     The Y key.
    Public Shared ReadOnly Property Y As Primitive = 68
    '
    ' Summary:
    '     The Z key.
    Public Shared ReadOnly Property Z As Primitive = 69
    '
    ' Summary:
    '     The left Windows logo key (Microsoft Natural Keyboard).
    Public Shared ReadOnly Property LWin As Primitive = 70
    '
    ' Summary:
    '     The right Windows logo key (Microsoft Natural Keyboard).
    Public Shared ReadOnly Property RWin As Primitive = 71
    '
    ' Summary:
    '     The Application key (Microsoft Natural Keyboard).
    Public Shared ReadOnly Property Apps As Primitive = 72
    '
    ' Summary:
    '     The Computer Sleep key.
    Public Shared ReadOnly Property Sleep As Primitive = 73
    '
    ' Summary:
    '     The 0 key on the numeric keypad.
    Public Shared ReadOnly Property NumPad0 As Primitive = 74
    '
    ' Summary:
    '     The 1 key on the numeric keypad.
    Public Shared ReadOnly Property NumPad1 As Primitive = 75
    '
    ' Summary:
    '     The 2 key on the numeric keypad.
    Public Shared ReadOnly Property NumPad2 As Primitive = 76
    '
    ' Summary:
    '     The 3 key on the numeric keypad.
    Public Shared ReadOnly Property NumPad3 As Primitive = 77
    '
    ' Summary:
    '     The 4 key on the numeric keypad.
    Public Shared ReadOnly Property NumPad4 As Primitive = 78
    '
    ' Summary:
    '     The 5 key on the numeric keypad.
    Public Shared ReadOnly Property NumPad5 As Primitive = 79
    '
    ' Summary:
    '     The 6 key on the numeric keypad.
    Public Shared ReadOnly Property NumPad6 As Primitive = 80
    '
    ' Summary:
    '     The 7 key on the numeric keypad.
    Public Shared ReadOnly Property NumPad7 As Primitive = 81
    '
    ' Summary:
    '     The 8 key on the numeric keypad.
    Public Shared ReadOnly Property NumPad8 As Primitive = 82
    '
    ' Summary:
    '     The 9 key on the numeric keypad.
    Public Shared ReadOnly Property NumPad9 As Primitive = 83
    '
    ' Summary:
    '     The Multiply key.
    Public Shared ReadOnly Property Multiply As Primitive = 84
    '
    ' Summary:
    '     The Add key.
    Public Shared ReadOnly Property Add As Primitive = 85
    '
    ' Summary:
    '     The Separator key.
    Public Shared ReadOnly Property Separator As Primitive = 86
    '
    ' Summary:
    '     The Subtract key.
    Public Shared ReadOnly Property Subtract As Primitive = 87
    '
    ' Summary:
    '     The Decimal key.
    Public Shared ReadOnly Property [Decimal] As Primitive = 88
    '
    ' Summary:
    '     The Divide key.
    Public Shared ReadOnly Property Divide As Primitive = 89
    '
    ' Summary:
    '     The F1 key.
    Public Shared ReadOnly Property F1 As Primitive = 90
    '
    ' Summary:
    '     The F2 key.
    Public Shared ReadOnly Property F2 As Primitive = 91
    '
    ' Summary:
    '     The F3 key.
    Public Shared ReadOnly Property F3 As Primitive = 92
    '
    ' Summary:
    '     The F4 key.
    Public Shared ReadOnly Property F4 As Primitive = 93
    '
    ' Summary:
    '     The F5 key.
    Public Shared ReadOnly Property F5 As Primitive = 94
    '
    ' Summary:
    '     The F6 key.
    Public Shared ReadOnly Property F6 As Primitive = 95
    '
    ' Summary:
    '     The F7 key.
    Public Shared ReadOnly Property F7 As Primitive = 96
    '
    ' Summary:
    '     The F8 key.
    Public Shared ReadOnly Property F8 As Primitive = 97
    '
    ' Summary:
    '     The F9 key.
    Public Shared ReadOnly Property F9 As Primitive = 98
    '
    ' Summary:
    '     The F10 key.
    Public Shared ReadOnly Property F10 As Primitive = 99
    '
    ' Summary:
    '     The F11 key.
    Public Shared ReadOnly Property F11 As Primitive = 100
    '
    ' Summary:
    '     The F12 key.
    Public Shared ReadOnly Property F12 As Primitive = 101
    '
    ' Summary:
    '     The F13 key.
    Public Shared ReadOnly Property F13 As Primitive = 102
    '
    ' Summary:
    '     The F14 key.
    Public Shared ReadOnly Property F14 As Primitive = 103
    '
    ' Summary:
    '     The F15 key.
    Public Shared ReadOnly Property F15 As Primitive = 104
    '
    ' Summary:
    '     The F16 key.
    Public Shared ReadOnly Property F16 As Primitive = 105
    '
    ' Summary:
    '     The F17 key.
    Public Shared ReadOnly Property F17 As Primitive = 106
    '
    ' Summary:
    '     The F18 key.
    Public Shared ReadOnly Property F18 As Primitive = 107
    '
    ' Summary:
    '     The F19 key.
    Public Shared ReadOnly Property F19 As Primitive = 108
    '
    ' Summary:
    '     The F20 key.
    Public Shared ReadOnly Property F20 As Primitive = 109
    '
    ' Summary:
    '     The F21 key.
    Public Shared ReadOnly Property F21 As Primitive = 110
    '
    ' Summary:
    '     The F22 key.
    Public Shared ReadOnly Property F22 As Primitive = 111
    '
    ' Summary:
    '     The F23 key.
    Public Shared ReadOnly Property F23 As Primitive = 112
    '
    ' Summary:
    '     The F24 key.
    Public Shared ReadOnly Property F24 As Primitive = 113
    '
    ' Summary:
    '     The Num Lock key.
    Public Shared ReadOnly Property NumLock As Primitive = 114
    '
    ' Summary:
    '     The Scroll Lock key.
    Public Shared ReadOnly Property Scroll As Primitive = 115
    '
    ' Summary:
    '     The left Shift key.
    Public Shared ReadOnly Property LeftShift As Primitive = 116
    '
    ' Summary:
    '     The right Shift key.
    Public Shared ReadOnly Property RightShift As Primitive = 117
    '
    ' Summary:
    '     The left CTRL key.
    Public Shared ReadOnly Property LeftCtrl As Primitive = 118
    '
    ' Summary:
    '     The right CTRL key.
    Public Shared ReadOnly Property RightCtrl As Primitive = 119
    '
    ' Summary:
    '     The left ALT key.
    Public Shared ReadOnly Property LeftAlt As Primitive = 120
    '
    ' Summary:
    '     The right ALT key.
    Public Shared ReadOnly Property RightAlt As Primitive = 121
    '
    ' Summary:
    '     The Browser Back key.
    Public Shared ReadOnly Property BrowserBack As Primitive = 122
    '
    ' Summary:
    '     The Browser Forward key.
    Public Shared ReadOnly Property BrowserForward As Primitive = 123
    '
    ' Summary:
    '     The Browser Refresh key.
    Public Shared ReadOnly Property BrowserRefresh As Primitive = 124
    '
    ' Summary:
    '     The Browser Stop key.
    Public Shared ReadOnly Property BrowserStop As Primitive = 125
    '
    ' Summary:
    '     The Browser Search key.
    Public Shared ReadOnly Property BrowserSearch As Primitive = 126
    '
    ' Summary:
    '     The Browser Favorites key.
    Public Shared ReadOnly Property BrowserFavorites As Primitive = 127
    '
    ' Summary:
    '     The Browser Home key.
    Public Shared ReadOnly Property BrowserHome As Primitive = 128
    '
    ' Summary:
    '     The Volume Mute key.
    Public Shared ReadOnly Property VolumeMute As Primitive = 129
    '
    ' Summary:
    '     The Volume Down key.
    Public Shared ReadOnly Property VolumeDown As Primitive = 130
    '
    ' Summary:
    '     The Volume Up key.
    Public Shared ReadOnly Property VolumeUp As Primitive = 131
    '
    ' Summary:
    '     The Media Next Track key.
    Public Shared ReadOnly Property MediaNextTrack As Primitive = 132
    '
    ' Summary:
    '     The Media Previous Track key.
    Public Shared ReadOnly Property MediaPreviousTrack As Primitive = 133
    '
    ' Summary:
    '     The Media Stop key.
    Public Shared ReadOnly Property MediaStop As Primitive = 134
    '
    ' Summary:
    '     The Media Play Pause key.
    Public Shared ReadOnly Property MediaPlayPause As Primitive = 135
    '
    ' Summary:
    '     The Launch Mail key.
    Public Shared ReadOnly Property LaunchMail As Primitive = 136
    '
    ' Summary:
    '     The Select Media key.
    Public Shared ReadOnly Property SelectMedia As Primitive = 137
    '
    ' Summary:
    '     The Launch Application1 key.
    Public Shared ReadOnly Property LaunchApplication1 As Primitive = 138
    '
    ' Summary:
    '     The Launch Application2 key.
    Public Shared ReadOnly Property LaunchApplication2 As Primitive = 139
    '
    ' Summary:
    '     The OEM 1 key.
    Public Shared ReadOnly Property Oem1 As Primitive = 140
    '
    ' Summary:
    '     The OEM Semicolon key.
    Public Shared ReadOnly Property OemSemicolon As Primitive = 140
    '
    ' Summary:
    '     The OEM Addition key.
    Public Shared ReadOnly Property OemPlus As Primitive = 141
    '
    ' Summary:
    '     The OEM Comma key.
    Public Shared ReadOnly Property OemComma As Primitive = 142
    '
    ' Summary:
    '     The OEM Minus key.
    Public Shared ReadOnly Property OemMinus As Primitive = 143
    '
    ' Summary:
    '     The OEM Period key.
    Public Shared ReadOnly Property OemPeriod As Primitive = 144
    '
    ' Summary:
    '     The OEM 2 key.
    Public Shared ReadOnly Property Oem2 As Primitive = 145
    '
    ' Summary:
    '     The OEM Question key.
    Public Shared ReadOnly Property OemQuestion As Primitive = 145
    '
    ' Summary:
    '     The OEM 3 key.
    Public Shared ReadOnly Property Oem3 As Primitive = 146
    '
    ' Summary:
    '     The OEM Tilde key.
    Public Shared ReadOnly Property OemTilde As Primitive = 146
    '
    ' Summary:
    '     The ABNT_C1 (Brazilian) key.
    Public Shared ReadOnly Property AbntC1 As Primitive = 147
    '
    ' Summary:
    '     The ABNT_C2 (Brazilian) key.
    Public Shared ReadOnly Property AbntC2 As Primitive = 148
    '
    ' Summary:
    '     The OEM 4 key.
    Public Shared ReadOnly Property Oem4 As Primitive = 149
    '
    ' Summary:
    '     The OEM Open Brackets key.
    Public Shared ReadOnly Property OemOpenBrackets As Primitive = 149
    '
    ' Summary:
    '     The OEM 5 key.
    Public Shared ReadOnly Property Oem5 As Primitive = 150
    '
    ' Summary:
    '     The OEM Pipe key.
    Public Shared ReadOnly Property OemPipe As Primitive = 150
    '
    ' Summary:
    '     The OEM 6 key.
    Public Shared ReadOnly Property Oem6 As Primitive = 151
    '
    ' Summary:
    '     The OEM Close Brackets key.
    Public Shared ReadOnly Property OemCloseBrackets As Primitive = 151
    '
    ' Summary:
    '     The OEM 7 key.
    Public Shared ReadOnly Property Oem7 As Primitive = 152
    '
    ' Summary:
    '     The OEM Quotes key.
    Public Shared ReadOnly Property OemQuotes As Primitive = 152
    '
    ' Summary:
    '     The OEM 8 key.
    Public Shared ReadOnly Property Oem8 As Primitive = 153
    '
    ' Summary:
    '     The OEM 102 key.
    Public Shared ReadOnly Property Oem102 As Primitive = 154
    '
    ' Summary:
    '     The OEM Backslash key.
    Public Shared ReadOnly Property OemBackslash As Primitive = 154
    '
    ' Summary:
    '     A special key masking the real key being processed by an IME.
    Public Shared ReadOnly Property ImeProcessed As Primitive = 155
    '
    ' Summary:
    '     A special key masking the real key being processed as a system key.
    Public Shared ReadOnly Property System As Primitive = 156
    '
    ' Summary:
    '     The OEM ATTN key.
    Public Shared ReadOnly Property OemAttn As Primitive = 157
    '
    ' Summary:
    '     The DBE_ALPHANUMERIC key.
    Public Shared ReadOnly Property DbeAlphanumeric As Primitive = 157
    '
    ' Summary:
    '     The OEM FINISH key.
    Public Shared ReadOnly Property OemFinish As Primitive = 158
    '
    ' Summary:
    '     The DBE_KATAKANA key.
    Public Shared ReadOnly Property DbeKatakana As Primitive = 158
    '
    ' Summary:
    '     The OEM COPY key.
    Public Shared ReadOnly Property OemCopy As Primitive = 159
    '
    ' Summary:
    '     The DBE_HIRAGANA key.
    Public Shared ReadOnly Property DbeHiragana As Primitive = 159
    '
    ' Summary:
    '     The OEM AUTO key.
    Public Shared ReadOnly Property OemAuto As Primitive = 160
    '
    ' Summary:
    '     The DBE_SBCSCHAR key.
    Public Shared ReadOnly Property DbeSbcsChar As Primitive = 160
    '
    ' Summary:
    '     The OEM ENLW key.
    Public Shared ReadOnly Property OemEnlw As Primitive = 161
    '
    ' Summary:
    '     The DBE_DBCSCHAR key.
    Public Shared ReadOnly Property DbeDbcsChar As Primitive = 161
    '
    ' Summary:
    '     The OEM BACKTAB key.
    Public Shared ReadOnly Property OemBackTab As Primitive = 162
    '
    ' Summary:
    '     The DBE_ROMAN key.
    Public Shared ReadOnly Property DbeRoman As Primitive = 162
    '
    ' Summary:
    '     The ATTN key.
    Public Shared ReadOnly Property Attn As Primitive = 163
    '
    ' Summary:
    '     The DBE_NOROMAN key.
    Public Shared ReadOnly Property DbeNoRoman As Primitive = 163
    '
    ' Summary:
    '     The CRSEL key.
    Public Shared ReadOnly Property CrSel As Primitive = 164
    '
    ' Summary:
    '     The DBE_ENTERWORDREGISTERMODE key.
    Public Shared ReadOnly Property DbeEnterWordRegisterMode As Primitive = 164
    '
    ' Summary:
    '     The EXSEL key.
    Public Shared ReadOnly Property ExSel As Primitive = 165
    '
    ' Summary:
    '     The DBE_ENTERIMECONFIGMODE key.
    Public Shared ReadOnly Property DbeEnterImeConfigureMode As Primitive = 165
    '
    ' Summary:
    '     The ERASE EOF key.
    Public Shared ReadOnly Property EraseEof As Primitive = 166
    '
    ' Summary:
    '     The DBE_FLUSHSTRING key.
    Public Shared ReadOnly Property DbeFlushString As Primitive = 166
    '
    ' Summary:
    '     The PLAY key.
    Public Shared ReadOnly Property Play As Primitive = 167
    '
    ' Summary:
    '     The DBE_CODEINPUT key.
    Public Shared ReadOnly Property DbeCodeInput As Primitive = 167
    '
    ' Summary:
    '     The ZOOM key.
    Public Shared ReadOnly Property Zoom As Primitive = 168
    '
    ' Summary:
    '     The DBE_NOCODEINPUT key.
    Public Shared ReadOnly Property DbeNoCodeInput As Primitive = 168
    '
    ' Summary:
    '     A constant reserved for future use.
    Public Shared ReadOnly Property NoName As Primitive = 169
    '
    ' Summary:
    '     The DBE_DETERMINESTRING key.
    Public Shared ReadOnly Property DbeDetermineString As Primitive = 169
    '
    ' Summary:
    '     The PA1 key.
    Public Shared ReadOnly Property Pa1 As Primitive = 170
    '
    ' Summary:
    '     The DBE_ENTERDLGCONVERSIONMODE key.
    Public Shared ReadOnly Property DbeEnterDialogConversionMode As Primitive = 170
    '
    ' Summary:
    '     The OEM Clear key.
    Public Shared ReadOnly Property OemClear As Primitive = 171
    '
    ' Summary:
    '     The key is used with another key to create a single combined character.
    Public Shared ReadOnly Property DeadCharProcessed As Primitive = 172

End Class
