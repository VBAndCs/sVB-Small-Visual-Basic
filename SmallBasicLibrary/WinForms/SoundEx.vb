Imports System.IO
Imports System.Media
Imports System.Threading
Imports System.Windows.Media
Imports Microsoft.SmallVisualBasic.Library
Imports Microsoft.SmallVisualBasic.Library.Internal

Namespace WinForms
    ''' <summary>
    ''' The Sound object provides operations that allow the playback of sounds.  Some sample sounds are provided along with the library.
    ''' The Sound object provides operations that allow the playback of sounds.  Some sample sounds are provided along with the library.
    ''' </summary>
    <SmallVisualBasicType>
    <HideFromIntellisense>
    Public NotInheritable Class SoundEx

        ''' <summary>Plays the current audio.</summary>
        <ExMethod>
        Public Shared Sub Play(filePath As Primitive)
            Sound.Play(filePath)
        End Sub

        ''' <summary>
        ''' Resumes playing the current audio from the position it was paused at.
        ''' If the file hasn't not played yet, this operation will play it from start.
        ''' </summary>
        <ExMethod>
        Public Shared Sub [Resume](filePath As Primitive)
            Sound.Resume(filePath)
        End Sub

        ''' <summary>
        ''' Plays the current audio and waits until it is finished playing.
        ''' If the file was already paused, this operation will resume from the position where the playback was paused.
        ''' </summary>
        <ExMethod>
        Public Shared Sub PlayAndWait(filePath As Primitive)
            Sound.PlayAndWait(filePath)
        End Sub

        ''' <summary>
        ''' Pauses playback of the current audio.  If the file was not already playing, this operation will not do anything.
        ''' </summary>
        <ExMethod>
        Public Shared Sub Pause(filePath As Primitive)
            Sound.Pause(filePath)
        End Sub

        ''' <summary>
        ''' Stops playback of the current audio.  If the file was not already playing, this operation will not do anything.
        ''' </summary>
        <ExMethod>
        Public Shared Sub [Stop](filePath As Primitive)
            Sound.Stop(filePath)
        End Sub


    End Class
End Namespace
