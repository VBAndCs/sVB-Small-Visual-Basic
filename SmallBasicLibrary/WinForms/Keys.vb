Imports Microsoft.SmallVisualBasic.Library

Namespace WinForms
    '
    '''<summary>Specifies the possible key values on a keyboard.</summary>
    <SmallVisualBasicType>
    Public NotInheritable Class Keys

        '''<summary>No key pressed.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property None As Primitive = 0

        '''<summary>The Cancel key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property Cancel As Primitive = 1

        '''<summary>The Backspace key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property Back As Primitive = 2

        '''<summary>The Tab key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property Tab As Primitive = 3

        '''<summary>The Linefeed key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property LineFeed As Primitive = 4

        '''<summary>The Clear key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property Clear As Primitive = 5

        '''<summary>The Return key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property [Return] As Primitive = 6

        '''<summary>The Enter key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property Enter As Primitive = 6

        '''<summary>The Pause key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property Pause As Primitive = 7

        '''<summary>The Caps Lock key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property Capital As Primitive = 8

        '''<summary>The Caps Lock key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property CapsLock As Primitive = 8

        '''<summary>The IME Kana mode key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property KanaMode As Primitive = 9

        '''<summary>The IME Hangul mode key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property HangulMode As Primitive = 9

        '''<summary>The IME Junja mode key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property JunjaMode As Primitive = 10

        '''<summary>The IME Final mode key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property FinalMode As Primitive = 11

        '''<summary>The IME Hanja mode key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property HanjaMode As Primitive = 12

        '''<summary>The IME Kanji mode key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property KanjiMode As Primitive = 12

        '''<summary>The ESC key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property Escape As Primitive = 13

        '''<summary>The IME Convert key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property ImeConvert As Primitive = 14

        '''<summary>The IME NonConvert key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property ImeNonConvert As Primitive = 15

        '''<summary>The IME Accept key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property ImeAccept As Primitive = 16

        '''<summary>The IME Mode change request.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property ImeModeChange As Primitive = 17

        '''<summary>The Spacebar key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property Space As Primitive = 18

        '''<summary>The Page Up key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property Prior As Primitive = 19

        '''<summary>The Page Up key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property PageUp As Primitive = 19

        '''<summary>The Page Down key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property [Next] As Primitive = 20

        '''<summary>The Page Down key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property PageDown As Primitive = 20

        '''<summary>The End key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property [End] As Primitive = 21

        '''<summary>The Home key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property Home As Primitive = 22

        '''<summary>The Left Arrow key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property Left As Primitive = 23

        '''<summary>The Up Arrow key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property Up As Primitive = 24

        '''<summary>The Right Arrow key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property Right As Primitive = 25

        '''<summary>The Down Arrow key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property Down As Primitive = 26

        '''<summary>The Select key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property [Select] As Primitive = 27

        '''<summary>The Print key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property Print As Primitive = 28

        '''<summary>The Execute key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property Execute As Primitive = 29

        '''<summary>The Print Screen key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property Snapshot As Primitive = 30

        '''<summary>The Print Screen key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property PrintScreen As Primitive = 30

        '''<summary>The Insert key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property Insert As Primitive = 31

        '''<summary>The Delete key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property Delete As Primitive = 32

        '''<summary>The Help key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property Help As Primitive = 33

        '''<summary>The 0 (zero) key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property D0 As Primitive = 34

        '''<summary>The 1 (one) key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property D1 As Primitive = 35

        '''<summary>The 2 key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property D2 As Primitive = 36

        '''<summary>The 3 key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property D3 As Primitive = 37

        '''<summary>The 4 key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property D4 As Primitive = 38

        '''<summary>The 5 key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property D5 As Primitive = 39

        '''<summary>The 6 key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property D6 As Primitive = 40

        '''<summary>The 7 key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property D7 As Primitive = 41

        '''<summary>The 8 key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property D8 As Primitive = 42

        '''<summary>The 9 key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property D9 As Primitive = 43

        '''<summary>The A key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property A As Primitive = 44

        '''<summary>The B key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property B As Primitive = 45

        '''<summary>The C key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property C As Primitive = 46

        '''<summary>The D key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property D As Primitive = 47

        '''<summary>The E key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property E As Primitive = 48

        '''<summary>The F key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property F As Primitive = 49

        '''<summary>The G key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property G As Primitive = 50

        '''<summary>The H key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property H As Primitive = 51

        '''<summary>The I key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property I As Primitive = 52

        '''<summary>The J key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property J As Primitive = 53

        '''<summary>The K key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property K As Primitive = 54

        '''<summary>The L key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property L As Primitive = 55

        '''<summary>The M key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property M As Primitive = 56

        '''<summary>The N key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property N As Primitive = 57

        '''<summary>The O key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property O As Primitive = 58

        '''<summary>The P key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property P As Primitive = 59

        '''<summary>The Q key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property Q As Primitive = 60

        '''<summary>The R key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property R As Primitive = 61

        '''<summary>The S key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property S As Primitive = 62

        '''<summary>The T key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property T As Primitive = 63

        '''<summary>The U key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property U As Primitive = 64

        '''<summary>The V key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property V As Primitive = 65

        '''<summary>The W key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property W As Primitive = 66

        '''<summary>The X key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property X As Primitive = 67

        '''<summary>The Y key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property Y As Primitive = 68

        '''<summary>The Z key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property Z As Primitive = 69

        '''<summary>The left Windows logo key (Microsoft Natural Keyboard).</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property LWin As Primitive = 70

        '''<summary>The right Windows logo key (Microsoft Natural Keyboard).</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property RWin As Primitive = 71

        '''<summary>The Application key (Microsoft Natural Keyboard).</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property Apps As Primitive = 72

        '''<summary>The Computer Sleep key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property Sleep As Primitive = 73

        '''<summary>The 0 key on the numeric keypad.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property NumPad0 As Primitive = 74

        '''<summary>The 1 key on the numeric keypad.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property NumPad1 As Primitive = 75

        '''<summary>The 2 key on the numeric keypad.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property NumPad2 As Primitive = 76

        '''<summary>The 3 key on the numeric keypad.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property NumPad3 As Primitive = 77

        '''<summary>The 4 key on the numeric keypad.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property NumPad4 As Primitive = 78

        '''<summary>The 5 key on the numeric keypad.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property NumPad5 As Primitive = 79

        '''<summary>The 6 key on the numeric keypad.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property NumPad6 As Primitive = 80

        '''<summary>The 7 key on the numeric keypad.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property NumPad7 As Primitive = 81

        '''<summary>The 8 key on the numeric keypad.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property NumPad8 As Primitive = 82

        '''<summary>The 9 key on the numeric keypad.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property NumPad9 As Primitive = 83

        '''<summary>The Multiply key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property Multiply As Primitive = 84

        '''<summary>The Add key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property Add As Primitive = 85

        '''<summary>The Separator key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property Separator As Primitive = 86

        '''<summary>The Subtract key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property Subtract As Primitive = 87

        '''<summary>The Decimal key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property [Decimal] As Primitive = 88

        '''<summary>The Divide key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property Divide As Primitive = 89

        '''<summary>The F1 key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property F1 As Primitive = 90

        '''<summary>The F2 key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property F2 As Primitive = 91

        '''<summary>The F3 key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property F3 As Primitive = 92

        '''<summary>The F4 key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property F4 As Primitive = 93

        '''<summary>The F5 key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property F5 As Primitive = 94

        '''<summary>The F6 key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property F6 As Primitive = 95

        '''<summary>The F7 key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property F7 As Primitive = 96

        '''<summary>The F8 key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property F8 As Primitive = 97

        '''<summary>The F9 key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property F9 As Primitive = 98

        '''<summary>The F10 key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property F10 As Primitive = 99

        '''<summary>The F11 key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property F11 As Primitive = 100

        '''<summary>The F12 key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property F12 As Primitive = 101

        '''<summary>The F13 key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property F13 As Primitive = 102

        '''<summary>The F14 key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property F14 As Primitive = 103

        '''<summary>The F15 key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property F15 As Primitive = 104

        '''<summary>The F16 key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property F16 As Primitive = 105

        '''<summary>The F17 key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property F17 As Primitive = 106

        '''<summary>The F18 key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property F18 As Primitive = 107

        '''<summary>The F19 key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property F19 As Primitive = 108

        '''<summary>The F20 key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property F20 As Primitive = 109

        '''<summary>The F21 key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property F21 As Primitive = 110

        '''<summary>The F22 key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property F22 As Primitive = 111

        '''<summary>The F23 key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property F23 As Primitive = 112

        '''<summary>The F24 key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property F24 As Primitive = 113

        '''<summary>The Num Lock key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property NumLock As Primitive = 114

        '''<summary>The Scroll Lock key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property Scroll As Primitive = 115

        '''<summary>The left Shift key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property LeftShift As Primitive = 116

        '''<summary>The right Shift key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property RightShift As Primitive = 117

        '''<summary>The left CTRL key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property LeftCtrl As Primitive = 118

        '''<summary>The right CTRL key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property RightCtrl As Primitive = 119

        '''<summary>The left ALT key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property LeftAlt As Primitive = 120

        '''<summary>The right ALT key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property RightAlt As Primitive = 121

        '''<summary>The Browser Back key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property BrowserBack As Primitive = 122

        '''<summary>The Browser Forward key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property BrowserForward As Primitive = 123

        '''<summary>The Browser Refresh key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property BrowserRefresh As Primitive = 124

        '''<summary>The Browser Stop key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property BrowserStop As Primitive = 125

        '''<summary>The Browser Search key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property BrowserSearch As Primitive = 126

        '''<summary>The Browser Favorites key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property BrowserFavorites As Primitive = 127

        '''<summary>The Browser Home key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property BrowserHome As Primitive = 128

        '''<summary>The Volume Mute key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property VolumeMute As Primitive = 129

        '''<summary>The Volume Down key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property VolumeDown As Primitive = 130

        '''<summary>The Volume Up key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property VolumeUp As Primitive = 131

        '''<summary>The Media Next Track key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property MediaNextTrack As Primitive = 132

        '''<summary>The Media Previous Track key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property MediaPreviousTrack As Primitive = 133

        '''<summary>The Media Stop key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property MediaStop As Primitive = 134

        '''<summary>The Media Play Pause key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property MediaPlayPause As Primitive = 135

        '''<summary>The Launch Mail key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property LaunchMail As Primitive = 136

        '''<summary>The Select Media key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property SelectMedia As Primitive = 137

        '''<summary>The Launch Application1 key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property LaunchApplication1 As Primitive = 138

        '''<summary>The Launch Application2 key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property LaunchApplication2 As Primitive = 139

        '''<summary>The OEM 1 key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property Oem1 As Primitive = 140

        '''<summary>The OEM Semicolon key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property OemSemicolon As Primitive = 140

        '''<summary>The OEM Addition key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property OemPlus As Primitive = 141

        '''<summary>The OEM Comma key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property OemComma As Primitive = 142

        '''<summary>The OEM Minus key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property OemMinus As Primitive = 143

        '''<summary>The OEM Period key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property OemPeriod As Primitive = 144

        '''<summary>The OEM 2 key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property Oem2 As Primitive = 145

        '''<summary>The OEM Question key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property OemQuestion As Primitive = 145

        '''<summary>The OEM 3 key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property Oem3 As Primitive = 146

        '''<summary>The OEM Tilde key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property OemTilde As Primitive = 146

        '''<summary>The ABNT_C1 (Brazilian) key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property AbntC1 As Primitive = 147

        '''<summary>The ABNT_C2 (Brazilian) key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property AbntC2 As Primitive = 148

        '''<summary>The OEM 4 key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property Oem4 As Primitive = 149

        '''<summary>The OEM Open Brackets key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property OemOpenBrackets As Primitive = 149

        '''<summary>The OEM 5 key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property Oem5 As Primitive = 150

        '''<summary>The OEM Pipe key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property OemPipe As Primitive = 150

        '''<summary>The OEM 6 key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property Oem6 As Primitive = 151

        '''<summary>The OEM Close Brackets key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property OemCloseBrackets As Primitive = 151

        '''<summary>The OEM 7 key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property Oem7 As Primitive = 152

        '''<summary>The OEM Quotes key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property OemQuotes As Primitive = 152

        '''<summary>The OEM 8 key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property Oem8 As Primitive = 153

        '''<summary>The OEM 102 key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property Oem102 As Primitive = 154

        '''<summary>The OEM Backslash key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property OemBackslash As Primitive = 154

        '''<summary>A special key masking the real key being processed by an IME.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property ImeProcessed As Primitive = 155

        '''<summary>A special key masking the real key being processed as a system key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property System As Primitive = 156

        '''<summary>The OEM ATTN key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property OemAttn As Primitive = 157

        '''<summary>The DBE_ALPHANUMERIC key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property DbeAlphanumeric As Primitive = 157

        '''<summary>The OEM FINISH key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property OemFinish As Primitive = 158

        '''<summary>The DBE_KATAKANA key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property DbeKatakana As Primitive = 158

        '''<summary>The OEM COPY key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property OemCopy As Primitive = 159

        '''<summary>The DBE_HIRAGANA key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property DbeHiragana As Primitive = 159

        '''<summary>The OEM AUTO key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property OemAuto As Primitive = 160

        '''<summary>The DBE_SBCSCHAR key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property DbeSbcsChar As Primitive = 160

        '''<summary>The OEM ENLW key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property OemEnlw As Primitive = 161

        '''<summary>The DBE_DBCSCHAR key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property DbeDbcsChar As Primitive = 161

        '''<summary>The OEM BACKTAB key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property OemBackTab As Primitive = 162

        '''<summary>The DBE_ROMAN key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property DbeRoman As Primitive = 162

        '''<summary>The ATTN key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property Attn As Primitive = 163

        '''<summary>The DBE_NOROMAN key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property DbeNoRoman As Primitive = 163

        '''<summary>The CRSEL key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property CrSel As Primitive = 164

        '''<summary>The DBE_ENTERWORDREGISTERMODE key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property DbeEnterWordRegisterMode As Primitive = 164

        '''<summary>The EXSEL key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property ExSel As Primitive = 165

        '''<summary>The DBE_ENTERIMECONFIGMODE key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property DbeEnterImeConfigureMode As Primitive = 165

        '''<summary>The ERASE EOF key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property EraseEof As Primitive = 166

        '''<summary>The DBE_FLUSHSTRING key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property DbeFlushString As Primitive = 166

        '''<summary>The PLAY key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property Play As Primitive = 167

        '''<summary>The DBE_CODEINPUT key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property DbeCodeInput As Primitive = 167

        '''<summary>The ZOOM key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property Zoom As Primitive = 168

        '''<summary>The DBE_NOCODEINPUT key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property DbeNoCodeInput As Primitive = 168

        '''<summary>A constant reserved for future use.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property NoName As Primitive = 169

        '''<summary>The DBE_DETERMINESTRING key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property DbeDetermineString As Primitive = 169

        '''<summary>The PA1 key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property Pa1 As Primitive = 170

        '''<summary>The DBE_ENTERDLGCONVERSIONMODE key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property DbeEnterDialogConversionMode As Primitive = 170

        '''<summary>The OEM Clear key.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property OemClear As Primitive = 171

        '''<summary>The key is used with another key to create a single combined character.</summary>
        <ReturnValueType(VariableType.Key)>
        Public Shared ReadOnly Property DeadCharProcessed As Primitive = 172

    End Class
End Namespace
