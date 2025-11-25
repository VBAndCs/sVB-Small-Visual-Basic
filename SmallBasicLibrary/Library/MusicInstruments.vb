Imports Microsoft.SmallVisualBasic.Library
Imports Microsoft.SmallVisualBasic.WinForms

''' <summary>
''' Provides General MIDI instrument IDs to be used with the Sound.MusicInstrument properties.
''' </summary>
<SmallVisualBasicType>
Public NotInheritable Class MusicInstruments


    ' -------------------- Piano (0–7) --------------------

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Acoustic Grand Piano instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>    
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property AcousticGrandPiano As Primitive = 0

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Bright Acoustic Piano instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property BrightAcousticPiano As Primitive = 1

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Electric Grand Piano instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property ElectricGrandPiano As Primitive = 2

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Honky-Tonk Piano instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property HonkyTonkPiano As Primitive = 3

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Electric Piano 1 instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property ElectricPiano1 As Primitive = 4

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Electric Piano 2 instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property ElectricPiano2 As Primitive = 5

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Harpsichord instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property Harpsichord As Primitive = 6

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Clavinet instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property Clavinet As Primitive = 7

    ' -------------------- Chromatic Percussion (8–15) --------------------

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Celesta instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property Celesta As Primitive = 8

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Glockenspiel instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property Glockenspiel As Primitive = 9

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Music Box instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property MusicBox As Primitive = 10

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Vibraphone instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property Vibraphone As Primitive = 11

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Marimba instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property Marimba As Primitive = 12

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Xylophone instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property Xylophone As Primitive = 13

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Tubular Bells instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property TubularBells As Primitive = 14

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Dulcimer instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property Dulcimer As Primitive = 15

    ' -------------------- Organ (16–23) --------------------

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Drawbar Organ instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property DrawbarOrgan As Primitive = 16

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Percussive Organ instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property PercussiveOrgan As Primitive = 17

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Rock Organ instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property RockOrgan As Primitive = 18

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Church Organ instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property ChurchOrgan As Primitive = 19

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Reed Organ instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property ReedOrgan As Primitive = 20

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Accordion instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property Accordion As Primitive = 21

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Harmonica instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property Harmonica As Primitive = 22

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Tango Accordion instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property TangoAccordion As Primitive = 23

    ' -------------------- Guitar (24–31) --------------------

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Acoustic Guitar (nylon) instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property AcousticGuitarNylon As Primitive = 24

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Acoustic Guitar (steel) instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property AcousticGuitarSteel As Primitive = 25

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Electric Guitar (jazz) instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property ElectricGuitarJazz As Primitive = 26

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Electric Guitar (clean) instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property ElectricGuitarClean As Primitive = 27

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Electric Guitar (muted) instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property ElectricGuitarMuted As Primitive = 28

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Overdriven Guitar instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property OverdrivenGuitar As Primitive = 29

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Distortion Guitar instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property DistortionGuitar As Primitive = 30

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Guitar harmonics instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property GuitarHarmonics As Primitive = 31

    ' -------------------- Bass (32–39) --------------------

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Acoustic Bass instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property AcousticBass As Primitive = 32

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Electric Bass (finger) instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property ElectricBassFinger As Primitive = 33

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Electric Bass (pick) instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property ElectricBassPick As Primitive = 34

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Fretless Bass instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property FretlessBass As Primitive = 35

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Slap Bass 1 instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property SlapBass1 As Primitive = 36

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Slap Bass 2 instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property SlapBass2 As Primitive = 37

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Synth Bass 1 instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property SynthBass1 As Primitive = 38

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Synth Bass 2 instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property SynthBass2 As Primitive = 39

    ' -------------------- Strings (40–47) --------------------

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Violin instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property Violin As Primitive = 40

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Viola instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property Viola As Primitive = 41

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Cello instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property Cello As Primitive = 42

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Contrabass instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property Contrabass As Primitive = 43

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Tremolo Strings instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property TremoloStrings As Primitive = 44

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Pizzicato Strings instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property PizzicatoStrings As Primitive = 45

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Orchestral Harp instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property OrchestralHarp As Primitive = 46

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Timpani instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property Timpani As Primitive = 47

    ' -------------------- Ensemble (48–55) --------------------

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the String Ensemble 1 instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property StringEnsemble1 As Primitive = 48

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the String Ensemble 2 instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property StringEnsemble2 As Primitive = 49

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Synth Strings 1 instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property SynthStrings1 As Primitive = 50

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Synth Strings 2 instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property SynthStrings2 As Primitive = 51

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Choir Aahs instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property ChoirAahs As Primitive = 52

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Voice Oohs instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property VoiceOohs As Primitive = 53

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Synth Voice instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property SynthVoice As Primitive = 54

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Orchestra Hit instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property OrchestraHit As Primitive = 55

    ' -------------------- Brass (56–63) --------------------

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Trumpet instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property Trumpet As Primitive = 56

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Trombone instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property Trombone As Primitive = 57

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Tuba instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property Tuba As Primitive = 58

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Muted Trumpet instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property MutedTrumpet As Primitive = 59

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the French Horn instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property FrenchHorn As Primitive = 60

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Brass Section instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property BrassSection As Primitive = 61

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Synth Brass 1 instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property SynthBrass1 As Primitive = 62

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Synth Brass 2 instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property SynthBrass2 As Primitive = 63

    ' -------------------- Reed (64–71) --------------------

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Soprano Sax instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property SopranoSax As Primitive = 64

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Alto Sax instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property AltoSax As Primitive = 65

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Tenor Sax instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property TenorSax As Primitive = 66

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Baritone Sax instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property BaritoneSax As Primitive = 67

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Oboe instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property Oboe As Primitive = 68

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the English Horn instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property EnglishHorn As Primitive = 69

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Bassoon instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property Bassoon As Primitive = 70

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Clarinet instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property Clarinet As Primitive = 71

    ' -------------------- Pipe (72–79) --------------------

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Piccolo instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property Piccolo As Primitive = 72

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Flute instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property Flute As Primitive = 73

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Recorder instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property Recorder As Primitive = 74

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Pan Flute instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property PanFlute As Primitive = 75

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Blown Bottle instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property BlownBottle As Primitive = 76

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Shakuhachi instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property Shakuhachi As Primitive = 77

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Whistle instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property Whistle As Primitive = 78

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Ocarina instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property Ocarina As Primitive = 79

    ' -------------------- Synth Lead (80–87) --------------------

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Lead 1 (square) instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property Lead1Square As Primitive = 80

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Lead 2 (sawtooth) instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property Lead2Sawtooth As Primitive = 81

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Lead 3 (calliope) instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property Lead3Calliope As Primitive = 82

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Lead 4 (chiff) instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property Lead4Chiff As Primitive = 83

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Lead 5 (charang) instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property Lead5Charang As Primitive = 84

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Lead 6 (voice) instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property Lead6Voice As Primitive = 85

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Lead 7 (fifths) instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property Lead7Fifths As Primitive = 86

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Lead 8 (bass + lead) instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property Lead8BassLead As Primitive = 87

    ' -------------------- Synth Pad (88–95) --------------------

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Pad 1 (new age) instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property Pad1NewAge As Primitive = 88

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Pad 2 (warm) instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property Pad2Warm As Primitive = 89

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Pad 3 (polysynth) instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property Pad3Polysynth As Primitive = 90

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Pad 4 (choir) instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property Pad4Choir As Primitive = 91

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Pad 5 (bowed) instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property Pad5Bowed As Primitive = 92

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Pad 6 (metallic) instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property Pad6Metallic As Primitive = 93

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Pad 7 (halo) instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property Pad7Halo As Primitive = 94

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Pad 8 (sweep) instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property Pad8Sweep As Primitive = 95

    ' -------------------- Synth Effects (96–103) --------------------

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the FX 1 (rain) instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property FX1Rain As Primitive = 96

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the FX 2 (soundtrack) instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property FX2Soundtrack As Primitive = 97

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the FX 3 (crystal) instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property FX3Crystal As Primitive = 98

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the FX 4 (atmosphere) instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property FX4Atmosphere As Primitive = 99

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the FX 5 (brightness) instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property FX5Brightness As Primitive = 100

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the FX 6 (goblins) instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property FX6Goblins As Primitive = 101

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the FX 7 (echoes) instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property FX7Echoes As Primitive = 102

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the FX 8 (sci-fi) instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property FX8SciFi As Primitive = 103

    ' -------------------- Ethnic (104–111) --------------------

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Sitar instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property Sitar As Primitive = 104

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Banjo instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property Banjo As Primitive = 105

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Shamisen instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property Shamisen As Primitive = 106

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Koto instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property Koto As Primitive = 107

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Kalimba instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property Kalimba As Primitive = 108

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Bag pipe instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property Bagpipe As Primitive = 109

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Fiddle instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property Fiddle As Primitive = 110

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Shanai instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property Shanai As Primitive = 111

    ' -------------------- Percussive (112–119) --------------------

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Tinkle Bell instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property TinkleBell As Primitive = 112

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Agogo instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property Agogo As Primitive = 113

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Steel Drums instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property SteelDrums As Primitive = 114

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Woodblock instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property Woodblock As Primitive = 115

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Taiko Drum instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property TaikoDrum As Primitive = 116

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Melodic Tom instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property MelodicTom As Primitive = 117

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Synth Drum instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property SynthDrum As Primitive = 118

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Reverse Cymbal instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property ReverseCymbal As Primitive = 119

    ' -------------------- Sound Effects (120–127) --------------------

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Guitar Fret Noise instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property GuitarFretNoise As Primitive = 120

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Breath Noise instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property BreathNoise As Primitive = 121

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Seashore instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property Seashore As Primitive = 122

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Bird Tweet instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property BirdTweet As Primitive = 123

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Telephone Ring instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property TelephoneRing As Primitive = 124

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Helicopter instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property Helicopter As Primitive = 125

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Applause instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property Applause As Primitive = 126

    ''' <summary>
    ''' Set this value to the Sound.MusicInstrument property, so you can use the Gunshot instrument to play the music notes you pass to the Sound.PlayMusic method.
    ''' </summary>
    <ReturnValueType(VariableType.MusicInstrument)>
    Public Shared ReadOnly Property Gunshot As Primitive = 127

End Class
