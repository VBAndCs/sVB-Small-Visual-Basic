Imports System.Collections
Imports System.Collections.Generic
Imports System.Collections.ObjectModel

Namespace Microsoft.Nautilus.Text
    Friend NotInheritable Class NormalizedTextChangeCollection
        Inherits ReadOnlyCollection(Of ITextChange)
        Implements INormalizedTextChangeCollection, IList(Of ITextChange), ICollection(Of ITextChange), IEnumerable(Of ITextChange), IEnumerable

        Public Sub New(changes As IList(Of TextChange))
            MyBase.New(Normalize(changes))
        End Sub

        Private Shared Function StableSort(changes As IList(Of TextChange)) As TextChange()
            Dim array As TextChange() = New TextChange(changes.Count - 1) {}
            changes.CopyTo(array, 0)
            For i As Integer = 0 To array.Length - 1 - 1
                Dim position1 As Integer = array(i).Position
                Dim num As Integer = i

                For j As Integer = i + 1 To array.Length - 1
                    If array(j).Position < position1 Then
                        position1 = array(j).Position
                        num = j
                    End If
                Next

                Dim textChange1 As TextChange = array(i)
                array(i) = array(num)
                array(num) = textChange1
            Next
            Return array
        End Function

        Private Shared Sub Catenate(ByRef left As String, right As String)
            If right.Length > 0 Then
                left = (If((left.Length = 0), right, (left & right)))
            End If
        End Sub

        Private Shared Function Normalize(changes As IList(Of TextChange)) As IList(Of ITextChange)
            If changes Is Nothing Then
                Throw New ArgumentNullException("changes")
            End If

            If changes.Count = 1 Then
                Return New ITextChange(0) {changes(0)}
            End If

            If changes.Count = 0 Then
                Return New ITextChange() {}
            End If

            Dim sortedChanges = StableSort(changes)
            Dim delta = 0
            Dim curIndex = 0
            Dim nextIndex = 1

            While nextIndex < sortedChanges.Length
                Dim curChange = sortedChanges(curIndex)
                Dim nextChange = sortedChanges(nextIndex)
                Dim distance = nextChange.Position - curChange.OldEnd

                If distance > 0 Then
                    curChange.Position += delta
                    delta += curChange.Delta
                    curIndex = nextIndex
                    nextIndex += 1
                    Continue While
                End If

                If curChange.Position = nextChange.Position AndAlso
                         curChange.Delta = 0 AndAlso nextChange.Delta = 0 Then
                    curChange._NewText = nextChange.NewText
                Else
                    Catenate(curChange._NewText, nextChange.NewText)
                    If distance = 0 Then
                        Catenate(curChange._OldText, nextChange.OldText)
                    ElseIf curChange.OldEnd < nextChange.OldEnd Then
                        Dim startIndex = curChange.OldEnd - nextChange.Position
                        Catenate(curChange._OldText, nextChange.OldText.Substring(startIndex))
                    End If
                End If

                sortedChanges(nextIndex) = Nothing
                nextIndex += 1
            End While

            sortedChanges(curIndex).Position += delta
            Dim normalizedChanges As New List(Of ITextChange)

            For i As Integer = 0 To sortedChanges.Length - 1
                If sortedChanges(i) IsNot Nothing Then
                    normalizedChanges.Add(sortedChanges(i))
                End If
            Next

            Return normalizedChanges
        End Function
    End Class
End Namespace
