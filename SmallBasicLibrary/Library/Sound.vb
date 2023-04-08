Imports System.IO
Imports System.Media
Imports System.Reflection
Imports System.Threading
Imports System.Windows.Media
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
            EnsureDeviceInit()
            If notes.IsArray Then
                For Each note In notes._arrayMap.Values
                    PlayNotes(note)
                Next
            Else
                PlayNotes(notes)
            End If
        End Sub

        ''' <summary>
        ''' <para>
        ''' Plays an audio file.  This could be an mp3 or wav or wma file.  Other file formats may or may not play depending on the audio codecs installed on the user's computer.
        ''' </para>
        ''' <para>
        ''' If the file was already paused, this operation will resume from the position where the playback was paused.
        ''' </para>
        ''' </summary>
        ''' <param name="filePath">
        ''' The path for the audio file.  This could either be a local file (e.g.: c:\music\track1.mp3) or a file on the network (e.g.: http://contoso.com/track01.wma).
        ''' </param>
        Public Shared Sub Play(filePath As Primitive)
            GetMediaPlayer(filePath)?.Play()
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
            Dim mediaPlayer1 As MediaPlayer = GetMediaPlayer(filePath)
            If mediaPlayer1 Is Nothing Then
                Return
            End If
            Dim autoResetEvent1 As New AutoResetEvent(initialState:=False)
            mediaPlayer1.Play()
            Dim num As Integer = 0
            While Not autoResetEvent1.WaitOne(200, exitContext:=False)
                If num > 30 AndAlso mediaPlayer1.Position = TimeSpan.Zero Then
                    autoResetEvent1.Set()
                End If
                If mediaPlayer1.NaturalDuration.HasTimeSpan AndAlso mediaPlayer1.Position >= mediaPlayer1.NaturalDuration.TimeSpan Then
                    autoResetEvent1.Set()
                End If
                num += 1
            End While
            mediaPlayer1.Stop()
        End Sub

        ''' <summary>
        ''' Pauses playback of an audio file.  If the file was not already playing, this operation will not do anything.
        ''' </summary>
        ''' <param name="filePath">
        ''' The path for the audio file.  This could either be a local file (e.g.: c:\music\track1.mp3) or a file on the network (e.g.: http://contoso.com/track01.wma).
        ''' </param>
        Public Shared Sub Pause(filePath As Primitive)
            GetMediaPlayer(filePath)?.Pause()
        End Sub

        ''' <summary>
        ''' Stops playback of an audio file.  If the file was not already playing, this operation will not do anything.
        ''' </summary>
        ''' <param name="filePath">
        ''' The path for the audio file.  This could either be a local file (e.g.: c:\music\track1.mp3) or a file on the network (e.g.: http://contoso.com/track01.wma).
        ''' </param>
        Public Shared Sub [Stop](filePath As Primitive)
            GetMediaPlayer(filePath)?.Stop()
        End Sub

        Private Shared Sub PlayStockSound(soundStream As Stream, sync As Boolean)
            Dim soundPlayer As New SoundPlayer(soundStream)
            If sync Then
                soundPlayer.PlaySync()
            Else
                soundPlayer.Play()
            End If
        End Sub

        Private Shared Function GetMediaPlayer(filePath As Primitive) As MediaPlayer
            If filePath.IsEmpty Then Return Nothing

            Try
                Dim uri1 As New Uri(filePath)
                Dim value As MediaPlayer = Nothing
                If Not _mediaPlayerMap.TryGetValue(uri1, value) Then
                    value = New MediaPlayer
                    _mediaPlayerMap(uri1) = value
                    value.Open(uri1)
                End If

                Return value
            Catch
            End Try

            Return Nothing
        End Function

        Private Shared Sub EnsureDeviceInit()
            If _midiOut = 0 Then
                NativeHelper.midiOutOpen(_midiOut, 0, 0, 0, 0UL)
                ChangeInstrument(0)
            End If
        End Sub

        Private Shared Sub PlayNotes(song As String)
            Dim num As Integer = 0
            song = song.ToUpperInvariant()
            Dim length1 As Integer = song.Length
            While num < song.Length
                Dim num2 As Integer = _defaultLength
                Dim c As Char = song(num)
                num += 1

                If Char.IsLetter(c) Then
                    Dim text As New String(c, 1)
                    If num < length1 Then
                        c = song(num)
                        If c = "#"c OrElse c = "+"c OrElse c = "-"c Then
                            text += New String(c, 1)
                            num += 1
                        End If
                        If num < length1 Then
                            c = song(num)
                            If Char.IsDigit(c) Then
                                num2 = AscW(c) - 48
                                num += 1
                            End If
                            If num < length1 Then
                                c = song(num)
                                If Char.IsDigit(c) Then
                                    num2 = num2 * 10 + (AscW(c) - 48)
                                    num += 1
                                End If
                            End If
                        End If
                    End If
                    If text(0) = "O"c Then
                        _octave = num2
                    Else
                        PlayNote(_octave, text, num2)
                    End If
                Else
                    Select Case c
                        Case ">"c
                            _octave = System.Math.Min(8, _octave + 1)
                        Case "<"c
                            _octave = System.Math.Max(0, _octave - 1)
                    End Select
                End If
            End While
        End Sub

        Private Shared Sub PlayNote(octave As Integer, note As String, length As Integer)
            Dim num As Double = 1600.0 / CDbl(length)
            Select Case note
                Case "P", "R"
                    Thread.Sleep(num)
                    Return
                Case "L"
                    _defaultLength = length
                    Return
            End Select

            Dim value As Integer = Nothing
            If Not _notes.TryGetValue(note, value) Then
                value = 0
            End If
            octave = System.Math.Min(System.Math.Max(0, octave), 8)
            Dim number As Integer = octave * 12 + value
            PlayNote(number)
            Thread.Sleep(num)
        End Sub

        Private Shared Sub PlayNote(number As Integer)
            Dim dwMsg As UInteger = BitConverter.ToUInt32(New Byte(3) {
                144,
                CByte(number),
                100,
                0
            }, 0)
            NativeHelper.midiOutShortMsg(_midiOut, dwMsg)
        End Sub

        Private Shared Sub ChangeInstrument(number As Integer)
            Dim dwMsg As UInteger = BitConverter.ToUInt32(New Byte(3) {
                192,
                CByte(number),
                0,
                0
            }, 0)
            NativeHelper.midiOutShortMsg(_midiOut, dwMsg)
        End Sub
    End Class
End Namespace
