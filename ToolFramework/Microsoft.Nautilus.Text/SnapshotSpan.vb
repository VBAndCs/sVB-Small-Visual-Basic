
Namespace Microsoft.Nautilus.Text
    Public Structure SnapshotSpan

        Public ReadOnly Property Snapshot As ITextSnapshot

        Public ReadOnly Property Span As Span

        Public ReadOnly Property Start As Integer
            Get
                Return _Span.Start
            End Get
        End Property

        Public ReadOnly Property [End] As Integer
            Get
                Return _Span.End
            End Get
        End Property

        Public ReadOnly Property Length As Integer
            Get
                Return _Span.Length
            End Get
        End Property

        Public ReadOnly Property IsEmpty As Boolean
            Get
                Return _Span.IsEmpty
            End Get
        End Property

        Public Sub New(snapshot As ITextSnapshot, span As Span)
            If snapshot Is Nothing Then
                Throw New ArgumentNullException("_snapshot")
            End If

            If span.End > snapshot.Length Then
                Throw New ArgumentOutOfRangeException("span")
            End If

            _Snapshot = snapshot
            _Span = span
        End Sub

        Public Sub New(_snapshot As ITextSnapshot, start1 As Integer, length1 As Integer)
            Me.New(_snapshot, New Span(start1, length1))
        End Sub

        Public Shared Widening Operator CType(snapshotSpan1 As SnapshotSpan) As Span
            Return snapshotSpan1._Span
        End Operator

        Public Function GetText() As String
            Return _Snapshot.GetText(_Span)
        End Function

        Public Function TranslateTo(targetSnapshot As ITextSnapshot, spanTrackingMode1 As SpanTrackingMode) As SnapshotSpan
            If targetSnapshot Is Nothing Then
                Throw New ArgumentNullException("targetSnapshot")
            End If

            If targetSnapshot.TextBuffer IsNot _Snapshot.TextBuffer Then
                Throw New ArgumentException("targetSnapshot belongs to the wrong TextBuffer")
            End If

            If targetSnapshot Is _Snapshot Then
                Return Me
            End If

            Dim textSpan As ITextSpan = _Snapshot.CreateTextSpan(_Span, spanTrackingMode1)
            Return New SnapshotSpan(targetSnapshot, textSpan.GetSpan(targetSnapshot))
        End Function

        Public Function Contains(position As Integer) As Boolean
            Return _Span.Contains(position)
        End Function

        Public Function Contains(simpleSpan As Span) As Boolean
            Return _Span.Contains(simpleSpan)
        End Function

        Public Function OverlapsWith(simpleSpan As Span) As Boolean
            Return _Span.OverlapsWith(simpleSpan)
        End Function

        Public Function Overlap(simpleSpan As Span) As Span?
            Return _Span.Overlap(simpleSpan)
        End Function

        Public Overrides Function ToString() As String
            Return _Span.ToString()
        End Function

        Public Overrides Function Equals(obj As Object) As Boolean
            If TypeOf obj Is SnapshotSpan Then
                Return CType(obj, SnapshotSpan) = Me
            End If
            Return False
        End Function

        Public Shared Operator =(ss1 As SnapshotSpan, ss2 As SnapshotSpan) As Boolean
            If ss1._Snapshot Is ss2._Snapshot Then
                Return ss1._Span = ss2._Span
            End If
            Return False
        End Operator

        Public Shared Operator <>(ss1 As SnapshotSpan, ss2 As SnapshotSpan) As Boolean
            Return Not (ss1 = ss2)
        End Operator
    End Structure
End Namespace
