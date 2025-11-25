Imports System.IO
Imports System.Media
Imports System.Threading
Imports System.Windows.Media
Imports System.Windows.Threading
Imports Microsoft.SmallVisualBasic.Library.Internal

Namespace Library
    ''' <summary>
    ''' The Sound object provides operations that allow the playback of sounds.  Some sample sounds are provided along with the library.
    ''' </summary>
    <SmallVisualBasicType>
    Public NotInheritable Class Sound
        Private Shared _mediaPlayerMap As New Dictionary(Of Uri, MediaPlayer)
        Private Shared _midiOut As Integer = 0
        Private Shared _octave As Integer = 4
        Private Shared _defaultLength As Integer = 4
        Private Shared _notes As New Dictionary(Of String, Integer) From {
            {"C", 0},
            {"C+", 1},
            {"C#", 1},
            {"D-", 1},
            {"D", 2},
            {"D+", 3},
            {"D#", 3},
            {"E-", 3},
            {"E", 4},
            {"F", 5},
            {"F+", 6},
            {"F#", 6},
            {"G-", 6},
            {"G", 7},
            {"G+", 8},
            {"G#", 8},
            {"A-", 8},
            {"A", 9},
            {"A+", 10},
            {"A#", 10},
            {"B-", 10},
            {"B", 11}
        }

        ''' <summary>
        ''' Plays the ding sound.
        ''' </summary>
        Public Shared Sub PlayDing()
            PlayStockSound(Nothing, sync:=False)
        End Sub

        ''' <summary>
        ''' Plays the beep beep sound.
        ''' </summary>
        Public Shared Sub PlayBeep()
            PlayStockSound(My.Resources.BeepBeep, sync:=False)
        End Sub

        ''' <summary>
        ''' Plays the beep beep sound and waits for it to finish.
        ''' </summary>
        Public Shared Sub PlayBeepAndWait()
            PlayStockSound(My.Resources.BeepBeep, sync:=True)
        End Sub

        ''' <summary>
        ''' Plays the Click Sound.
        ''' </summary>
        Public Shared Sub PlayClick()
            PlayStockSound(My.Resources.Click, sync:=False)
        End Sub

        ''' <summary>
        ''' Plays the Click Sound and waits for it to finish.
        ''' </summary>
        Public Shared Sub PlayClickAndWait()
            PlayStockSound(My.Resources.Click, sync:=True)
        End Sub

        ''' <summary>
        ''' Plays the Chime Sound.
        ''' </summary>
        Public Shared Sub PlayChime()
            PlayStockSound(My.Resources.Chime, sync:=False)
        End Sub

        ''' <summary>
        ''' Plays the Chime Sound and waits for it to finish.
        ''' </summary>
        Public Shared Sub PlayChimeAndWait()
            PlayStockSound(My.Resources.Chime, sync:=True)
        End Sub

        ''' <summary>
        ''' Plays the Chimes Sound.
        ''' </summary>
        Public Shared Sub PlayChimes()
            PlayStockSound(My.Resources.Pause, sync:=False)
        End Sub

        ''' <summary>
        ''' Plays the Chimes Sound and waits for it to finish.
        ''' </summary>
        Public Shared Sub PlayChimesAndWait()
            PlayStockSound(My.Resources.Pause, sync:=True)
        End Sub

        ''' <summary>
        ''' Plays the Bell Ring Sound.
        ''' </summary>
        Public Shared Sub PlayBellRing()
            PlayStockSound(My.Resources.BellRing, sync:=False)
        End Sub

        ''' <summary>
        ''' Plays the Bell Ring Sound and waits for it to finish.
        ''' </summary>
        Public Shared Sub PlayBellRingAndWait()
            PlayStockSound(My.Resources.BellRing, sync:=True)
        End Sub

        ''' <summary>
        ''' Plays musical notes without deylaing the execution of the program code.
        ''' This allows you to play more than one note a the same time and control the delay between them using the Programe.Delay.
        ''' </summary>
        ''' <param name="notes">
        ''' A set of musical notes to play.  The format is a subset of the Music Macro Language supported by QBasic.
        ''' </param>
        Public Shared Sub PlayMusicAsync(notes As Primitive)
            Dim channel = _musicChannel
            Dim instrument = _musicInstrument
            Dim volume = _volume
            Dim keepTonesOn = _keepTonesOn
            Task.Run(Sub() DoPlayMusic(notes, instrument, channel, volume, keepTonesOn))
        End Sub

        ''' <summary>
        ''' Plays musical notes.
        ''' </summary>
        ''' <param name="notes">
        ''' A set of musical notes to play.  The format is a subset of the Music Macro Language supported by QBasic.
        ''' </param>
        ''' <example>
        ''' <code>
        ''' Sound.PlayMusic("O5 C8 C8 G8 G8 A8 A8 G4 F8 F8 E8 E8 D8 D8 C4")
        ''' </code>
        ''' </example>
        Public Shared Sub PlayMusic(notes As Primitive)
            DoPlayMusic(notes, _musicInstrument, _musicChannel, _volume * 127 / 100, _keepTonesOn)
        End Sub

        Private Shared Sub DoPlayMusic(
                notes As Primitive,
                instrument As Byte,
                channel As Byte,
                volume As Byte,
                keepTonesOn As Boolean)

            Dim defaultLength = _defaultLength
            Dim globalLength = _defaultLength

            EnsureDeviceInit(instrument, channel)
            volume *= 1.27

            If notes.IsArray Then
                For Each note In notes.ArrayMap.Values
                    PlayNotes(note, channel, volume, keepTonesOn, defaultLength, globalLength)
                Next
            Else
                PlayNotes(notes, channel, volume, keepTonesOn, defaultLength, globalLength)
            End If
        End Sub

        ''' <summary>
        ''' Loads an audio file, so you can call the Play method to play it later without suffering from any initial delay.
        ''' </summary>
        ''' <param name="filePath">
        ''' The path for the audio file.  This could either be a local file (e.g.: c:\music\track1.mp3) or a file on the network (e.g.: http://contoso.com/track01.wma).
        ''' You can load mp3, wav and wma files. Other file formats may or may not be valid depending on the audio codecs installed on the user's computer.
        ''' </param>
        ''' <returns>If the audio file is loaded correctly, this method will return the file path, so you can send it to the Play, Pause and Stop methods, otherwise this method will return an empty string.</returns>
        <WinForms.ReturnValueType(VariableType.Sound)>
        Public Shared Function Load(filePath As Primitive) As Primitive
            SmallBasicApplication.Invoke(
                Sub()
                    Dim mp = GetMediaPlayer(filePath)

                    If mp Is Nothing Then
                        Load = New Primitive("")
                    Else
                        Load = filePath
                    End If
                End Sub)
        End Function


        ''' <summary>
        ''' Plays the given audio file. 
        ''' If the file is currently played or paused, this operation will replay it from the start. Use the Resume method to continue playing the paused file.
        ''' </summary>
        ''' <param name="filePath">
        ''' The path for the audio file. This could either be a local file (e.g.: c:\music\track1.mp3) or a file on the network (e.g.: http://contoso.com/track01.wma).
        ''' ''' The file type could be an mp3 or wav or wma file. Other file formats may or may not play depending on the audio codecs installed on the user's computer.
        ''' </param>
        Public Shared Sub Play(filePath As Primitive)
            SmallBasicApplication.Invoke(
                Sub()
                    Dim mp = GetMediaPlayer(filePath)
                    If mp IsNot Nothing Then
                        mp.Stop()
                        mp.Play()
                    End If

                End Sub)
        End Sub

        ''' <summary>
        ''' Resumes playing the given audio from the position it was paused at.
        ''' If the file hasn't not played yet, this operation will play it from start.
        ''' </summary>
        ''' <param name="filePath">
        ''' The path for the audio file. This could either be a local file (e.g.: c:\music\track1.mp3) or a file on the network (e.g.: http://contoso.com/track01.wma).
        ''' The file type could be an mp3 or wav or wma file. Other file formats may or may not play depending on the audio codecs installed on the user's computer.
        ''' </param>
        Public Shared Sub [Resume](filePath As Primitive)
            SmallBasicApplication.Invoke(
                Sub()
                    Dim mp = GetMediaPlayer(filePath)
                    If mp IsNot Nothing Then mp.Play()
                End Sub)
        End Sub

        ''' <summary>
        ''' <para>
        ''' Plays an audio file and waits until it is finished playing.  This could be an mp3 or wav or wma file. Other file formats may or may not play depending on the audio codecs installed on the user's computer.
        ''' </para>
        ''' <para>
        ''' If the file was already paused, this operation will resume from the position where the playback was paused.
        ''' </para>
        ''' </summary>
        ''' <param name="filePath">
        ''' The path for the audio file.  This could either be a local file (e.g.: c:\music\track1.mp3) or a file on the network (e.g.: http://contoso.com/track01.wma).
        ''' </param>
        Public Shared Sub PlayAndWait(filePath As Primitive)
            SmallBasicApplication.Invoke(
                Sub()
                    Dim mediaPlayer = GetMediaPlayer(filePath)
                    If mediaPlayer Is Nothing Then Return

                    Dim autoResetEvent1 As New AutoResetEvent(initialState:=False)
                    mediaPlayer.Play()
                    Dim n As Integer = 0
                    While Not autoResetEvent1.WaitOne(200, exitContext:=False)
                        If n > 30 AndAlso mediaPlayer.Position = TimeSpan.Zero Then
                            autoResetEvent1.Set()
                        End If

                        If mediaPlayer.NaturalDuration.HasTimeSpan AndAlso mediaPlayer.Position >= mediaPlayer.NaturalDuration.TimeSpan Then
                            autoResetEvent1.Set()
                        End If

                        n += 1
                    End While
                    mediaPlayer.Stop()
                End Sub)
        End Sub

        ''' <summary>
        ''' Pauses playback of an audio file.  If the file was not already playing, this operation will not do anything.
        ''' </summary>
        ''' <param name="filePath">
        ''' The path for the audio file.  This could either be a local file (e.g.: c:\music\track1.mp3) or a file on the network (e.g.: http://contoso.com/track01.wma).
        ''' </param>
        Public Shared Sub Pause(filePath As Primitive)
            SmallBasicApplication.Invoke(
                Sub()
                    GetMediaPlayer(filePath)?.Pause()
                End Sub)
        End Sub

        ''' <summary>
        ''' Stops playback of an audio file.  If the file was not already playing, this operation will not do anything.
        ''' </summary>
        ''' <param name="filePath">
        ''' The path for the audio file.  This could either be a local file (e.g.: c:\music\track1.mp3) or a file on the network (e.g.: http://contoso.com/track01.wma).
        ''' </param>
        Public Shared Sub [Stop](filePath As Primitive)
            SmallBasicApplication.Invoke(
                Sub()
                    GetMediaPlayer(filePath)?.Stop()
                End Sub)
        End Sub

        Private Shared Sub PlayStockSound(soundStream As Stream, sync As Boolean)
            Dim soundPlayer As New SoundPlayer(soundStream)
            If sync Then
                soundPlayer.PlaySync()
            Else
                soundPlayer.Play()
            End If
        End Sub

        Private Shared _volume As Byte = 50

        ''' <summary>
        ''' Gets or sets the volume of the played sound files. Valid values are between 0 and 100. The default value is 50.
        ''' Note that changing this value will only affect the files you play afterwards.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Property Volume As Primitive
            Get
                Return _volume
            End Get

            Set(value As Primitive)
                Dim v As Integer = value.AsDecimal
                Select Case v
                    Case < 0
                        _volume = 0
                    Case > 100
                        _volume = 100
                    Case Else
                        _volume = v
                End Select
            End Set
        End Property

        Private Shared Function GetMediaPlayer(filePath As Primitive) As MediaPlayer
            If filePath.IsEmpty Then Return Nothing

            SmallBasicApplication.Invoke(
                Sub()
                    Try
                        Dim localFileName As String = Network.GetLocalFile(filePath)
                        Dim uri1 As New Uri(localFileName, UriKind.RelativeOrAbsolute)
                        Dim mp As MediaPlayer = Nothing

                        If Not _mediaPlayerMap.TryGetValue(uri1, mp) Then
                            mp = New MediaPlayer()
                            _mediaPlayerMap(uri1) = mp
                            mp.Open(uri1)
                        End If

                        mp.Volume = _volume / 100
                        GetMediaPlayer = mp

                    Catch ex As Exception
                        ReportError("Error in loading sound. Make sure that the filee name is valid "" " & filePath & """", ex)
                    End Try
                End Sub)

        End Function

        Private Shared Sub EnsureDeviceInit(instrument As Byte, channel As Byte)
            If _midiOut = 0 Then
                NativeHelper.midiOutOpen(_midiOut, 0, 0, 0, 0UL)
            End If

            ChangeInstrument(instrument, channel)
        End Sub

        Private Shared Sub PlayNotes(
                song As String,
                channel As Byte,
                volume As Byte,
                keepTonesOn As Boolean,
                ByRef defaultLength As Integer,
                ByRef globalLength As Integer)

            Dim i = 0
            song = song.ToUpperInvariant().Trim()
            Dim songLength = song.Length

            If songLength = 0 Then
                PlayNote(60, channel, 0)
                Return
            End If

            Do While i < songLength
                If globalLength <> _defaultLength Then
                    defaultLength = _defaultLength
                    globalLength = _defaultLength
                End If

                Dim length = defaultLength
                Dim c As Char = song(i)
                i += 1

                If Char.IsLetter(c) Then
                    Dim note As New String(c, 1)
                    If i < songLength Then
                        c = song(i)
                        If c = "#"c OrElse c = "+"c OrElse c = "-"c Then
                            note += New String(c, 1)
                            i += 1
                        End If

                        If i < songLength Then
                            c = song(i)
                            If Char.IsDigit(c) Then
                                length = AscW(c) - 48
                                i += 1
                            End If

                            If i < songLength Then
                                c = song(i)
                                If Char.IsDigit(c) Then
                                    length = length * 10 + (AscW(c) - 48)
                                    i += 1
                                End If
                            End If
                        End If
                    End If

                    If note(0) = "O"c Then
                        _octave = length
                    Else
                        PlayNote(_octave, note, length, channel, volume, keepTonesOn, defaultLength)
                    End If

                Else
                    Select Case c
                        Case ">"c
                            _octave = System.Math.Min(8, _octave + 1)
                        Case "<"c
                            _octave = System.Math.Max(0, _octave - 1)
                    End Select
                End If
            Loop
        End Sub

        Private Shared _keepTonesOn As Boolean = True

        ''' <summary>
        ''' Determines whether tones (played with the PlayMuseic and PlayMusicAsync methods) are allowed to sustain naturally or are forcibly terminated after their duration. 
        ''' The default value is True, which is suitable for some instruments like piano and guitar.
        ''' You should use False with some instruments (like sax and harmonica) to avoid noise.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Boolean)>
        Public Shared Property KeepTonesOn As Primitive
            Get
                Return New Primitive(_keepTonesOn)
            End Get

            Set(value As Primitive)
                _keepTonesOn = CBool(value)
            End Set
        End Property

        Private Shared Sub PlayNote(
                           octave As Integer,
                           note As String,
                           length As Integer,
                           channel As Byte,
                           volume As Byte,
                           keepTonesOn As Boolean,
                           ByRef defaultLength As Integer)

            Dim delay As Double = 1600 / length
            Select Case note
                Case "P", "R"
                    Thread.Sleep(delay)
                    Return
                Case "S"
                    defaultLength = length
                    Return
                Case "L"
                    _defaultLength = length
                    defaultLength = length
                    Return
            End Select

            Dim value As Integer = Nothing
            If Not _notes.TryGetValue(note, value) Then
                value = 0
            End If

            octave = System.Math.Min(System.Math.Max(0, octave), 8)
            Dim number = octave * 12 + value
            PlayNote(number, channel, volume)
            Thread.Sleep(delay)
            If Not keepTonesOn Then StopNote(number, channel)
        End Sub

        Private Shared Sub PlayNote(noteNumber As Integer, channel As Byte, volume As Byte)
            Dim dwMsg = BitConverter.ToUInt32(
                New Byte(3) {
                    144 + channel,
                    CByte(noteNumber),
                    volume,
                    0},
                0)
            NativeHelper.midiOutShortMsg(_midiOut, dwMsg)
        End Sub

        Private Shared Sub StopNote(noteNumber As Integer, channel As Byte)
            Dim dwMsg = BitConverter.ToUInt32(
                New Byte(3) {
                    128 + channel,
                    CByte(noteNumber),
                    0,   ' velocity = 0
                    0},
                0)
            NativeHelper.midiOutShortMsg(_midiOut, dwMsg)
        End Sub

        Private Shared _musicChannel As Byte = 0

        ''' <summary>
        ''' Gets or sets the music channel used by the PlayMusic method.
        ''' Valid values are between 0 and 15. Use 9 for percussion instruments (beats).
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Property MusicChannel As Primitive
            Get
                Return _musicChannel
            End Get

            Set(value As Primitive)
                Dim v = value.AsDecimal
                If v < 0 Then
                    _musicChannel = 0
                ElseIf v > 15 Then
                    _musicChannel = 15
                Else
                    _musicChannel = v
                End If
            End Set
        End Property

        Private Shared _musicInstrument As Byte = 0

        ''' <summary>
        ''' Gets or sets the music instrument that will be used by the PlayMusic method.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.MusicInstrument)>
        Public Shared Property MusicInstrument As Primitive
            Get
                Return _musicInstrument
            End Get

            Set(value As Primitive)
                Dim v = value.AsDecimal
                If v < 0 Then
                    _musicInstrument = 0
                ElseIf v > 127 Then
                    _musicInstrument = 127
                Else
                    _musicInstrument = v
                End If
            End Set
        End Property

        Private Shared Sub ChangeInstrument(instrumentNumber As Byte, channel As Byte)
            Dim dwMsg = BitConverter.ToUInt32(
                New Byte(3) {
                    CByte(192 + channel),
                    instrumentNumber,
                    0, 0},
                0)
            NativeHelper.midiOutShortMsg(_midiOut, dwMsg)
        End Sub
    End Class
End Namespace
