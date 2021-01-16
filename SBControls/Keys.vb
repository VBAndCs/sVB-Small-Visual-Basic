Imports Microsoft.SmallBasic.Library

'
' Summary:
'     Specifies the possible key values on a keyboard.
<SmallBasicType>
Public Module Keys

    '
    ' Summary:
    '     No key pressed.
    Public ReadOnly Property None As Primitive = 0
    '
    ' Summary:
    '     The Cancel key.
    Public ReadOnly Property Cancel As Primitive = 1
    '
    ' Summary:
    '     The Backspace key.
    Public ReadOnly Property Back As Primitive = 2
    '
    ' Summary:
    '     The Tab key.
    Public ReadOnly Property Tab As Primitive = 3
    '
    ' Summary:
    '     The Linefeed key.
    Public ReadOnly Property LineFeed As Primitive = 4
    '
    ' Summary:
    '     The Clear key.
    Public ReadOnly Property Clear As Primitive = 5
    '
    ' Summary:
    '     The Return key.
    Public ReadOnly Property [Return] As Primitive = 6
    '
    ' Summary:
    '     The Enter key.
    Public ReadOnly Property Enter As Primitive = 6
    '
    ' Summary:
    '     The Pause key.
    Public ReadOnly Property Pause As Primitive = 7
    '
    ' Summary:
    '     The Caps Lock key.
    Public ReadOnly Property Capital As Primitive = 8
    '
    ' Summary:
    '     The Caps Lock key.
    Public ReadOnly Property CapsLock As Primitive = 8
    '
    ' Summary:
    '     The IME Kana mode key.
    Public ReadOnly Property KanaMode As Primitive = 9
    '
    ' Summary:
    '     The IME Hangul mode key.
    Public ReadOnly Property HangulMode As Primitive = 9
    '
    ' Summary:
    '     The IME Junja mode key.
    Public ReadOnly Property JunjaMode As Primitive = 10
    '
    ' Summary:
    '     The IME Final mode key.
    Public ReadOnly Property FinalMode As Primitive = 11
    '
    ' Summary:
    '     The IME Hanja mode key.
    Public ReadOnly Property HanjaMode As Primitive = 12
    '
    ' Summary:
    '     The IME Kanji mode key.
    Public ReadOnly Property KanjiMode As Primitive = 12
    '
    ' Summary:
    '     The ESC key.
    Public ReadOnly Property Escape As Primitive = 13
    '
    ' Summary:
    '     The IME Convert key.
    Public ReadOnly Property ImeConvert As Primitive = 14
    '
    ' Summary:
    '     The IME NonConvert key.
    Public ReadOnly Property ImeNonConvert As Primitive = 15
    '
    ' Summary:
    '     The IME Accept key.
    Public ReadOnly Property ImeAccept As Primitive = 16
    '
    ' Summary:
    '     The IME Mode change request.
    Public ReadOnly Property ImeModeChange As Primitive = 17
    '
    ' Summary:
    '     The Spacebar key.
    Public ReadOnly Property Space As Primitive = 18
    '
    ' Summary:
    '     The Page Up key.
    Public ReadOnly Property Prior As Primitive = 19
    '
    ' Summary:
    '     The Page Up key.
    Public ReadOnly Property PageUp As Primitive = 19
    '
    ' Summary:
    '     The Page Down key.
    Public ReadOnly Property [Next] As Primitive = 20
    '
    ' Summary:
    '     The Page Down key.
    Public ReadOnly Property PageDown As Primitive = 20
    '
    ' Summary:
    '     The End key.
    Public ReadOnly Property [End] As Primitive = 21
    '
    ' Summary:
    '     The Home key.
    Public ReadOnly Property Home As Primitive = 22
    '
    ' Summary:
    '     The Left Arrow key.
    Public ReadOnly Property Left As Primitive = 23
    '
    ' Summary:
    '     The Up Arrow key.
    Public ReadOnly Property Up As Primitive = 24
    '
    ' Summary:
    '     The Right Arrow key.
    Public ReadOnly Property Right As Primitive = 25
    '
    ' Summary:
    '     The Down Arrow key.
    Public ReadOnly Property Down As Primitive = 26
    '
    ' Summary:
    '     The Select key.
    Public ReadOnly Property [Select] As Primitive = 27
    '
    ' Summary:
    '     The Print key.
    Public ReadOnly Property Print As Primitive = 28
    '
    ' Summary:
    '     The Execute key.
    Public ReadOnly Property Execute As Primitive = 29
    '
    ' Summary:
    '     The Print Screen key.
    Public ReadOnly Property Snapshot As Primitive = 30
    '
    ' Summary:
    '     The Print Screen key.
    Public ReadOnly Property PrintScreen As Primitive = 30
    '
    ' Summary:
    '     The Insert key.
    Public ReadOnly Property Insert As Primitive = 31
    '
    ' Summary:
    '     The Delete key.
    Public ReadOnly Property Delete As Primitive = 32
    '
    ' Summary:
    '     The Help key.
    Public ReadOnly Property Help As Primitive = 33
    '
    ' Summary:
    '     The 0 (zero) key.
    Public ReadOnly Property D0 As Primitive = 34
    '
    ' Summary:
    '     The 1 (one) key.
    Public ReadOnly Property D1 As Primitive = 35
    '
    ' Summary:
    '     The 2 key.
    Public ReadOnly Property D2 As Primitive = 36
    '
    ' Summary:
    '     The 3 key.
    Public ReadOnly Property D3 As Primitive = 37
    '
    ' Summary:
    '     The 4 key.
    Public ReadOnly Property D4 As Primitive = 38
    '
    ' Summary:
    '     The 5 key.
    Public ReadOnly Property D5 As Primitive = 39
    '
    ' Summary:
    '     The 6 key.
    Public ReadOnly Property D6 As Primitive = 40
    '
    ' Summary:
    '     The 7 key.
    Public ReadOnly Property D7 As Primitive = 41
    '
    ' Summary:
    '     The 8 key.
    Public ReadOnly Property D8 As Primitive = 42
    '
    ' Summary:
    '     The 9 key.
    Public ReadOnly Property D9 As Primitive = 43
    '
    ' Summary:
    '     The A key.
    Public ReadOnly Property A As Primitive = 44
    '
    ' Summary:
    '     The B key.
    Public ReadOnly Property B As Primitive = 45
    '
    ' Summary:
    '     The C key.
    Public ReadOnly Property C As Primitive = 46
    '
    ' Summary:
    '     The D key.
    Public ReadOnly Property D As Primitive = 47
    '
    ' Summary:
    '     The E key.
    Public ReadOnly Property E As Primitive = 48
    '
    ' Summary:
    '     The F key.
    Public ReadOnly Property F As Primitive = 49
    '
    ' Summary:
    '     The G key.
    Public ReadOnly Property G As Primitive = 50
    '
    ' Summary:
    '     The H key.
    Public ReadOnly Property H As Primitive = 51
    '
    ' Summary:
    '     The I key.
    Public ReadOnly Property I As Primitive = 52
    '
    ' Summary:
    '     The J key.
    Public ReadOnly Property J As Primitive = 53
    '
    ' Summary:
    '     The K key.
    Public ReadOnly Property K As Primitive = 54
    '
    ' Summary:
    '     The L key.
    Public ReadOnly Property L As Primitive = 55
    '
    ' Summary:
    '     The M key.
    Public ReadOnly Property M As Primitive = 56
    '
    ' Summary:
    '     The N key.
    Public ReadOnly Property N As Primitive = 57
    '
    ' Summary:
    '     The O key.
    Public ReadOnly Property O As Primitive = 58
    '
    ' Summary:
    '     The P key.
    Public ReadOnly Property P As Primitive = 59
    '
    ' Summary:
    '     The Q key.
    Public ReadOnly Property Q As Primitive = 60
    '
    ' Summary:
    '     The R key.
    Public ReadOnly Property R As Primitive = 61
    '
    ' Summary:
    '     The S key.
    Public ReadOnly Property S As Primitive = 62
    '
    ' Summary:
    '     The T key.
    Public ReadOnly Property T As Primitive = 63
    '
    ' Summary:
    '     The U key.
    Public ReadOnly Property U As Primitive = 64
    '
    ' Summary:
    '     The V key.
    Public ReadOnly Property V As Primitive = 65
    '
    ' Summary:
    '     The W key.
    Public ReadOnly Property W As Primitive = 66
    '
    ' Summary:
    '     The X key.
    Public ReadOnly Property X As Primitive = 67
    '
    ' Summary:
    '     The Y key.
    Public ReadOnly Property Y As Primitive = 68
    '
    ' Summary:
    '     The Z key.
    Public ReadOnly Property Z As Primitive = 69
    '
    ' Summary:
    '     The left Windows logo key (Microsoft Natural Keyboard).
    Public ReadOnly Property LWin As Primitive = 70
    '
    ' Summary:
    '     The right Windows logo key (Microsoft Natural Keyboard).
    Public ReadOnly Property RWin As Primitive = 71
    '
    ' Summary:
    '     The Application key (Microsoft Natural Keyboard).
    Public ReadOnly Property Apps As Primitive = 72
    '
    ' Summary:
    '     The Computer Sleep key.
    Public ReadOnly Property Sleep As Primitive = 73
    '
    ' Summary:
    '     The 0 key on the numeric keypad.
    Public ReadOnly Property NumPad0 As Primitive = 74
    '
    ' Summary:
    '     The 1 key on the numeric keypad.
    Public ReadOnly Property NumPad1 As Primitive = 75
    '
    ' Summary:
    '     The 2 key on the numeric keypad.
    Public ReadOnly Property NumPad2 As Primitive = 76
    '
    ' Summary:
    '     The 3 key on the numeric keypad.
    Public ReadOnly Property NumPad3 As Primitive = 77
    '
    ' Summary:
    '     The 4 key on the numeric keypad.
    Public ReadOnly Property NumPad4 As Primitive = 78
    '
    ' Summary:
    '     The 5 key on the numeric keypad.
    Public ReadOnly Property NumPad5 As Primitive = 79
    '
    ' Summary:
    '     The 6 key on the numeric keypad.
    Public ReadOnly Property NumPad6 As Primitive = 80
    '
    ' Summary:
    '     The 7 key on the numeric keypad.
    Public ReadOnly Property NumPad7 As Primitive = 81
    '
    ' Summary:
    '     The 8 key on the numeric keypad.
    Public ReadOnly Property NumPad8 As Primitive = 82
    '
    ' Summary:
    '     The 9 key on the numeric keypad.
    Public ReadOnly Property NumPad9 As Primitive = 83
    '
    ' Summary:
    '     The Multiply key.
    Public ReadOnly Property Multiply As Primitive = 84
    '
    ' Summary:
    '     The Add key.
    Public ReadOnly Property Add As Primitive = 85
    '
    ' Summary:
    '     The Separator key.
    Public ReadOnly Property Separator As Primitive = 86
    '
    ' Summary:
    '     The Subtract key.
    Public ReadOnly Property Subtract As Primitive = 87
    '
    ' Summary:
    '     The Decimal key.
    Public ReadOnly Property [Decimal] As Primitive = 88
    '
    ' Summary:
    '     The Divide key.
    Public ReadOnly Property Divide As Primitive = 89
    '
    ' Summary:
    '     The F1 key.
    Public ReadOnly Property F1 As Primitive = 90
    '
    ' Summary:
    '     The F2 key.
    Public ReadOnly Property F2 As Primitive = 91
    '
    ' Summary:
    '     The F3 key.
    Public ReadOnly Property F3 As Primitive = 92
    '
    ' Summary:
    '     The F4 key.
    Public ReadOnly Property F4 As Primitive = 93
    '
    ' Summary:
    '     The F5 key.
    Public ReadOnly Property F5 As Primitive = 94
    '
    ' Summary:
    '     The F6 key.
    Public ReadOnly Property F6 As Primitive = 95
    '
    ' Summary:
    '     The F7 key.
    Public ReadOnly Property F7 As Primitive = 96
    '
    ' Summary:
    '     The F8 key.
    Public ReadOnly Property F8 As Primitive = 97
    '
    ' Summary:
    '     The F9 key.
    Public ReadOnly Property F9 As Primitive = 98
    '
    ' Summary:
    '     The F10 key.
    Public ReadOnly Property F10 As Primitive = 99
    '
    ' Summary:
    '     The F11 key.
    Public ReadOnly Property F11 As Primitive = 100
    '
    ' Summary:
    '     The F12 key.
    Public ReadOnly Property F12 As Primitive = 101
    '
    ' Summary:
    '     The F13 key.
    Public ReadOnly Property F13 As Primitive = 102
    '
    ' Summary:
    '     The F14 key.
    Public ReadOnly Property F14 As Primitive = 103
    '
    ' Summary:
    '     The F15 key.
    Public ReadOnly Property F15 As Primitive = 104
    '
    ' Summary:
    '     The F16 key.
    Public ReadOnly Property F16 As Primitive = 105
    '
    ' Summary:
    '     The F17 key.
    Public ReadOnly Property F17 As Primitive = 106
    '
    ' Summary:
    '     The F18 key.
    Public ReadOnly Property F18 As Primitive = 107
    '
    ' Summary:
    '     The F19 key.
    Public ReadOnly Property F19 As Primitive = 108
    '
    ' Summary:
    '     The F20 key.
    Public ReadOnly Property F20 As Primitive = 109
    '
    ' Summary:
    '     The F21 key.
    Public ReadOnly Property F21 As Primitive = 110
    '
    ' Summary:
    '     The F22 key.
    Public ReadOnly Property F22 As Primitive = 111
    '
    ' Summary:
    '     The F23 key.
    Public ReadOnly Property F23 As Primitive = 112
    '
    ' Summary:
    '     The F24 key.
    Public ReadOnly Property F24 As Primitive = 113
    '
    ' Summary:
    '     The Num Lock key.
    Public ReadOnly Property NumLock As Primitive = 114
    '
    ' Summary:
    '     The Scroll Lock key.
    Public ReadOnly Property Scroll As Primitive = 115
    '
    ' Summary:
    '     The left Shift key.
    Public ReadOnly Property LeftShift As Primitive = 116
    '
    ' Summary:
    '     The right Shift key.
    Public ReadOnly Property RightShift As Primitive = 117
    '
    ' Summary:
    '     The left CTRL key.
    Public ReadOnly Property LeftCtrl As Primitive = 118
    '
    ' Summary:
    '     The right CTRL key.
    Public ReadOnly Property RightCtrl As Primitive = 119
    '
    ' Summary:
    '     The left ALT key.
    Public ReadOnly Property LeftAlt As Primitive = 120
    '
    ' Summary:
    '     The right ALT key.
    Public ReadOnly Property RightAlt As Primitive = 121
    '
    ' Summary:
    '     The Browser Back key.
    Public ReadOnly Property BrowserBack As Primitive = 122
    '
    ' Summary:
    '     The Browser Forward key.
    Public ReadOnly Property BrowserForward As Primitive = 123
    '
    ' Summary:
    '     The Browser Refresh key.
    Public ReadOnly Property BrowserRefresh As Primitive = 124
    '
    ' Summary:
    '     The Browser Stop key.
    Public ReadOnly Property BrowserStop As Primitive = 125
    '
    ' Summary:
    '     The Browser Search key.
    Public ReadOnly Property BrowserSearch As Primitive = 126
    '
    ' Summary:
    '     The Browser Favorites key.
    Public ReadOnly Property BrowserFavorites As Primitive = 127
    '
    ' Summary:
    '     The Browser Home key.
    Public ReadOnly Property BrowserHome As Primitive = 128
    '
    ' Summary:
    '     The Volume Mute key.
    Public ReadOnly Property VolumeMute As Primitive = 129
    '
    ' Summary:
    '     The Volume Down key.
    Public ReadOnly Property VolumeDown As Primitive = 130
    '
    ' Summary:
    '     The Volume Up key.
    Public ReadOnly Property VolumeUp As Primitive = 131
    '
    ' Summary:
    '     The Media Next Track key.
    Public ReadOnly Property MediaNextTrack As Primitive = 132
    '
    ' Summary:
    '     The Media Previous Track key.
    Public ReadOnly Property MediaPreviousTrack As Primitive = 133
    '
    ' Summary:
    '     The Media Stop key.
    Public ReadOnly Property MediaStop As Primitive = 134
    '
    ' Summary:
    '     The Media Play Pause key.
    Public ReadOnly Property MediaPlayPause As Primitive = 135
    '
    ' Summary:
    '     The Launch Mail key.
    Public ReadOnly Property LaunchMail As Primitive = 136
    '
    ' Summary:
    '     The Select Media key.
    Public ReadOnly Property SelectMedia As Primitive = 137
    '
    ' Summary:
    '     The Launch Application1 key.
    Public ReadOnly Property LaunchApplication1 As Primitive = 138
    '
    ' Summary:
    '     The Launch Application2 key.
    Public ReadOnly Property LaunchApplication2 As Primitive = 139
    '
    ' Summary:
    '     The OEM 1 key.
    Public ReadOnly Property Oem1 As Primitive = 140
    '
    ' Summary:
    '     The OEM Semicolon key.
    Public ReadOnly Property OemSemicolon As Primitive = 140
    '
    ' Summary:
    '     The OEM Addition key.
    Public ReadOnly Property OemPlus As Primitive = 141
    '
    ' Summary:
    '     The OEM Comma key.
    Public ReadOnly Property OemComma As Primitive = 142
    '
    ' Summary:
    '     The OEM Minus key.
    Public ReadOnly Property OemMinus As Primitive = 143
    '
    ' Summary:
    '     The OEM Period key.
    Public ReadOnly Property OemPeriod As Primitive = 144
    '
    ' Summary:
    '     The OEM 2 key.
    Public ReadOnly Property Oem2 As Primitive = 145
    '
    ' Summary:
    '     The OEM Question key.
    Public ReadOnly Property OemQuestion As Primitive = 145
    '
    ' Summary:
    '     The OEM 3 key.
    Public ReadOnly Property Oem3 As Primitive = 146
    '
    ' Summary:
    '     The OEM Tilde key.
    Public ReadOnly Property OemTilde As Primitive = 146
    '
    ' Summary:
    '     The ABNT_C1 (Brazilian) key.
    Public ReadOnly Property AbntC1 As Primitive = 147
    '
    ' Summary:
    '     The ABNT_C2 (Brazilian) key.
    Public ReadOnly Property AbntC2 As Primitive = 148
    '
    ' Summary:
    '     The OEM 4 key.
    Public ReadOnly Property Oem4 As Primitive = 149
    '
    ' Summary:
    '     The OEM Open Brackets key.
    Public ReadOnly Property OemOpenBrackets As Primitive = 149
    '
    ' Summary:
    '     The OEM 5 key.
    Public ReadOnly Property Oem5 As Primitive = 150
    '
    ' Summary:
    '     The OEM Pipe key.
    Public ReadOnly Property OemPipe As Primitive = 150
    '
    ' Summary:
    '     The OEM 6 key.
    Public ReadOnly Property Oem6 As Primitive = 151
    '
    ' Summary:
    '     The OEM Close Brackets key.
    Public ReadOnly Property OemCloseBrackets As Primitive = 151
    '
    ' Summary:
    '     The OEM 7 key.
    Public ReadOnly Property Oem7 As Primitive = 152
    '
    ' Summary:
    '     The OEM Quotes key.
    Public ReadOnly Property OemQuotes As Primitive = 152
    '
    ' Summary:
    '     The OEM 8 key.
    Public ReadOnly Property Oem8 As Primitive = 153
    '
    ' Summary:
    '     The OEM 102 key.
    Public ReadOnly Property Oem102 As Primitive = 154
    '
    ' Summary:
    '     The OEM Backslash key.
    Public ReadOnly Property OemBackslash As Primitive = 154
    '
    ' Summary:
    '     A special key masking the real key being processed by an IME.
    Public ReadOnly Property ImeProcessed As Primitive = 155
    '
    ' Summary:
    '     A special key masking the real key being processed as a system key.
    Public ReadOnly Property System As Primitive = 156
    '
    ' Summary:
    '     The OEM ATTN key.
    Public ReadOnly Property OemAttn As Primitive = 157
    '
    ' Summary:
    '     The DBE_ALPHANUMERIC key.
    Public ReadOnly Property DbeAlphanumeric As Primitive = 157
    '
    ' Summary:
    '     The OEM FINISH key.
    Public ReadOnly Property OemFinish As Primitive = 158
    '
    ' Summary:
    '     The DBE_KATAKANA key.
    Public ReadOnly Property DbeKatakana As Primitive = 158
    '
    ' Summary:
    '     The OEM COPY key.
    Public ReadOnly Property OemCopy As Primitive = 159
    '
    ' Summary:
    '     The DBE_HIRAGANA key.
    Public ReadOnly Property DbeHiragana As Primitive = 159
    '
    ' Summary:
    '     The OEM AUTO key.
    Public ReadOnly Property OemAuto As Primitive = 160
    '
    ' Summary:
    '     The DBE_SBCSCHAR key.
    Public ReadOnly Property DbeSbcsChar As Primitive = 160
    '
    ' Summary:
    '     The OEM ENLW key.
    Public ReadOnly Property OemEnlw As Primitive = 161
    '
    ' Summary:
    '     The DBE_DBCSCHAR key.
    Public ReadOnly Property DbeDbcsChar As Primitive = 161
    '
    ' Summary:
    '     The OEM BACKTAB key.
    Public ReadOnly Property OemBackTab As Primitive = 162
    '
    ' Summary:
    '     The DBE_ROMAN key.
    Public ReadOnly Property DbeRoman As Primitive = 162
    '
    ' Summary:
    '     The ATTN key.
    Public ReadOnly Property Attn As Primitive = 163
    '
    ' Summary:
    '     The DBE_NOROMAN key.
    Public ReadOnly Property DbeNoRoman As Primitive = 163
    '
    ' Summary:
    '     The CRSEL key.
    Public ReadOnly Property CrSel As Primitive = 164
    '
    ' Summary:
    '     The DBE_ENTERWORDREGISTERMODE key.
    Public ReadOnly Property DbeEnterWordRegisterMode As Primitive = 164
    '
    ' Summary:
    '     The EXSEL key.
    Public ReadOnly Property ExSel As Primitive = 165
    '
    ' Summary:
    '     The DBE_ENTERIMECONFIGMODE key.
    Public ReadOnly Property DbeEnterImeConfigureMode As Primitive = 165
    '
    ' Summary:
    '     The ERASE EOF key.
    Public ReadOnly Property EraseEof As Primitive = 166
    '
    ' Summary:
    '     The DBE_FLUSHSTRING key.
    Public ReadOnly Property DbeFlushString As Primitive = 166
    '
    ' Summary:
    '     The PLAY key.
    Public ReadOnly Property Play As Primitive = 167
    '
    ' Summary:
    '     The DBE_CODEINPUT key.
    Public ReadOnly Property DbeCodeInput As Primitive = 167
    '
    ' Summary:
    '     The ZOOM key.
    Public ReadOnly Property Zoom As Primitive = 168
    '
    ' Summary:
    '     The DBE_NOCODEINPUT key.
    Public ReadOnly Property DbeNoCodeInput As Primitive = 168
    '
    ' Summary:
    '     A constant reserved for future use.
    Public ReadOnly Property NoName As Primitive = 169
    '
    ' Summary:
    '     The DBE_DETERMINESTRING key.
    Public ReadOnly Property DbeDetermineString As Primitive = 169
    '
    ' Summary:
    '     The PA1 key.
    Public ReadOnly Property Pa1 As Primitive = 170
    '
    ' Summary:
    '     The DBE_ENTERDLGCONVERSIONMODE key.
    Public ReadOnly Property DbeEnterDialogConversionMode As Primitive = 170
    '
    ' Summary:
    '     The OEM Clear key.
    Public ReadOnly Property OemClear As Primitive = 171
    '
    ' Summary:
    '     The key is used with another key to create a single combined character.
    Public ReadOnly Property DeadCharProcessed As Primitive = 172

End Module
