
Namespace Microsoft.Nautilus.Text
    Public Structure SnapshotPoint

        Public ReadOnly Property Position As Integer

        Public ReadOnly Property Snapshot As ITextSnapshot

        Public Sub New(snapshot As ITextSnapshot, position As Integer)
            If snapshot Is Nothing Then
                Throw New ArgumentNullException("_Snapshot")
            End If

            If position < 0 OrElse position > snapshot.Length Then
                Throw New ArgumentOutOfRangeException("_Position")
            End If

            _Snapshot = snapshot
            _Position = position
        End Sub

        Public Shared Widening Operator CType(snapshotPoint1 As SnapshotPoint) As Integer
            Return snapshotPoint1.Position
        End Operator

        Public Shared Function ToInt32(snapshotPoint1 As SnapshotPoint) As Integer
            Return snapshotPoint1.Position
        End Function

        Public Function GetContainingLine() As ITextSnapshotLine
            Return _Snapshot.GetLineFromPosition(_Position)
        End Function

        Public Function GetChar() As Char
            Return _Snapshot(_Position)
        End Function

        Public Function TranslateTo(targetSnapshot As ITextSnapshot, trackingMode1 As TrackingMode) As SnapshotPoint
            If targetSnapshot Is Nothing Then
                Throw New ArgumentNullException("targetSnapshot")
            End If

            If targetSnapshot.TextBuffer IsNot _Snapshot.TextBuffer Then
                Throw New ArgumentException("targetSnapshot belongs to the wrong TextBuffer")
            End If

            If targetSnapshot Is _Snapshot Then Return Me

            Dim textPoint = _Snapshot.CreateTextPoint(_Position, trackingMode1)
            Return New SnapshotPoint(targetSnapshot, textPoint.GetPosition(targetSnapshot))
        End Function

        Public Overrides Function Equals(obj As Object) As Boolean
            If TypeOf obj IsNot SnapshotPoint Then Return False
            Return CType(obj, SnapshotPoint) = Me
        End Function

        Public Shared Operator =(sp1 As SnapshotPoint, sp2 As SnapshotPoint) As Boolean
            Return sp1.Snapshot Is sp2.Snapshot AndAlso sp1.Position = sp2.Position
        End Operator

        Public Shared Operator <>(sp1 As SnapshotPoint, sp2 As SnapshotPoint) As Boolean
            Return Not (sp1 = sp2)
        End Operator
    End Structure
End Namespace
