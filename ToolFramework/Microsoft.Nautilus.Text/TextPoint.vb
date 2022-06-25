Imports System.Collections.Generic

Namespace Microsoft.Nautilus.Text
    Public Class TextPoint
        Implements ITextPoint

        Private Class VersionPosition

            Public ReadOnly Property Version As TextVersion

            Public ReadOnly Property Position As Integer

            Public Sub New(version As TextVersion, position As Integer)
                _Version = version
                _Position = position
            End Sub
        End Class

        Private version As TextVersion
        Private position As Integer
        Private cachedPosition As VersionPosition

        Public ReadOnly Property TextBuffer As ITextBuffer Implements ITextPoint.TextBuffer
            Get
                Return version.TextBuffer
            End Get
        End Property

        Public ReadOnly Property TrackingMode As TrackingMode Implements ITextPoint.TrackingMode

        Public Sub New(snapshot As ITextSnapshot, position As Integer, trackingMode As TrackingMode)
            If snapshot Is Nothing Then
                Throw New ArgumentNullException("snapshot")
            End If

            If (position < 0) Or (position > snapshot.Length) Then
                Throw New ArgumentOutOfRangeException("position")
            End If

            version = snapshot.Version
            Me.position = position
            Me.TrackingMode = trackingMode
            cachedPosition = New VersionPosition(version, Me.position)
        End Sub

        Public Function GetPosition(snapshot As ITextSnapshot) As Integer Implements ITextPoint.GetPosition
            If snapshot Is Nothing Then
                Throw New ArgumentNullException("snapshot")
            End If

            If snapshot.TextBuffer IsNot TextBuffer Then
                Throw New ArgumentException("The specified ITextSnapshot doesn't belong to the correct text buffer.")
            End If

            Dim versionPosition1 As VersionPosition = cachedPosition
            Dim result As Integer
            If snapshot.Version > versionPosition1.Version Then
                result = MapPositionForwardInTime(versionPosition1.Position, _TrackingMode, snapshot.Version, versionPosition1.Version)
                cachedPosition = New VersionPosition(snapshot.Version, result)

            ElseIf snapshot.Version Is versionPosition1.Version Then
                result = versionPosition1.Position

            ElseIf Not (snapshot.Version > version) Then
                result = (If((snapshot.Version IsNot version), MapPositionBackwardsInTime(position, _TrackingMode, snapshot.Version, version), position))

            Else
                result = MapPositionForwardInTime(position, _TrackingMode, snapshot.Version, version)
                cachedPosition = New VersionPosition(snapshot.Version, result)
            End If
            Return result
        End Function

        Public Function GetPoint(snapshot As ITextSnapshot) As SnapshotPoint Implements ITextPoint.GetPoint
            Return New SnapshotPoint(snapshot, GetPosition(snapshot))
        End Function

        Public Function GetCharacter(snapshot As ITextSnapshot) As Char Implements ITextPoint.GetCharacter
            Return GetPoint(snapshot).GetChar()
        End Function

        Friend Shared Function MapPositionForwardInTime(position As Integer, _TrackingMode As TrackingMode, targetVersion As TextVersion, currentVersion As TextVersion) As Integer
            While currentVersion IsNot targetVersion
                If currentVersion.Changes IsNot Nothing Then
                    For Each change As ITextChange In currentVersion.Changes
                        If _TrackingMode = TrackingMode.Positive Then
                            If change.Position <= position Then
                                position = (If((change.OldEnd > position), change.NewEnd, (position + change.Delta)))
                            End If
                        ElseIf change.Position < position Then
                            position = (If((change.OldEnd >= position), change.Position, (position + change.Delta)))
                        End If
                    Next
                End If

                currentVersion = currentVersion.Next
            End While
            Return position
        End Function

        Friend Shared Function MapPositionBackwardsInTime(position As Integer, _TrackingMode As TrackingMode, targetVersion As TextVersion, currentVersion As TextVersion) As Integer
            Dim stack1 As New Stack(Of IList(Of ITextChange))
            Dim textVersion1 As TextVersion = targetVersion
            While textVersion1 IsNot currentVersion
                stack1.Push(textVersion1.Changes)
                textVersion1 = textVersion1.Next
            End While

            While stack1.Count > 0
                Dim list As IList(Of ITextChange) = stack1.Pop()
                If list Is Nothing Then
                    Continue While
                End If
                For num As Integer = list.Count - 1 To 0 Step -1
                    Dim textChange As ITextChange = list(num)
                    If _TrackingMode = TrackingMode.Positive Then
                        If textChange.Position <= position Then
                            position = (If((textChange.NewEnd > position), textChange.OldEnd, (position - textChange.Delta)))
                        End If
                    ElseIf textChange.Position < position Then
                        position = (If((textChange.NewEnd >= position), textChange.Position, (position - textChange.Delta)))
                    End If
                Next
            End While

            Return position
        End Function
    End Class
End Namespace
