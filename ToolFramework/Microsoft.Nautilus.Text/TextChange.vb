Imports System.Collections.Generic
Imports System.Text

Namespace Microsoft.Nautilus.Text
    Friend Class TextChange
        Implements ITextChange

        Public Property Position As Integer Implements ITextChange.Position

        Public ReadOnly Property Delta As Integer Implements ITextChange.Delta
            Get
                Return _newText.Length - _oldText.Length
            End Get
        End Property

        Public ReadOnly Property OldEnd As Integer Implements ITextChange.OldEnd
            Get
                Return _position + _oldText.Length
            End Get
        End Property

        Public ReadOnly Property NewEnd As Integer Implements ITextChange.NewEnd
            Get
                Return _position + _newText.Length
            End Get
        End Property

        Friend _OldText As String
        Public ReadOnly Property OldText As String Implements ITextChange.OldText
            Get
                Return _OldText
            End Get
        End Property

        Friend _NewText As String

        Public ReadOnly Property NewText As String Implements ITextChange.NewText
            Get
                Return _NewText
            End Get
        End Property

        Public ReadOnly Property NewLength As Integer Implements ITextChange.NewLength
            Get
                Return _newText.Length
            End Get
        End Property

        Public ReadOnly Property OldLength As Integer Implements ITextChange.OldLength
            Get
                Return _oldText.Length
            End Get
        End Property

        Public Sub New(position As Integer, oldText As String, newText As String)
            If position < 0 Then
                Throw New ArgumentOutOfRangeException("position")
            End If

            _Position = position
            _OldText = oldText
            _NewText = newText

            If _OldText Is Nothing Then _OldText = ""
            If _NewText Is Nothing Then _NewText = ""
        End Sub

        Public Shared Function Merge(normalizedChanges As IList(Of ITextChange), snapshot As ITextSnapshot) As ITextChange
            If normalizedChanges Is Nothing Then
                Throw New ArgumentNullException("normalizedChanges")
            End If
            If normalizedChanges.Count < 1 Then
                Throw New ArgumentOutOfRangeException("normalizedChanges.Count")
            End If
            If snapshot Is Nothing Then
                Throw New ArgumentNullException("snapshot")
            End If
            Dim position1 As Integer = normalizedChanges(0).Position
            Dim stringBuilder1 As New StringBuilder
            Dim stringBuilder2 As New StringBuilder
            For i As Integer = 0 To normalizedChanges.Count - 1
                Dim textChange1 As ITextChange = normalizedChanges(i)
                stringBuilder1.Append(textChange1.OldText)
                stringBuilder2.Append(textChange1.NewText)
                If i + 1 < normalizedChanges.Count Then
                    Dim length1 As Integer = normalizedChanges(i + 1).Position - normalizedChanges(i).NewEnd
                    Dim text As String = snapshot.GetText(normalizedChanges(i).NewEnd, length1)
                    stringBuilder1.Append(text)
                    stringBuilder2.Append(text)
                End If
            Next
            Return New TextChange(position1, stringBuilder1.ToString(), stringBuilder2.ToString())
        End Function
    End Class
End Namespace
