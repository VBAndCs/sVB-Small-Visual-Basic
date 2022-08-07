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

            Dim array As TextChange() = StableSort(changes)
            Dim num As Integer = 0
            Dim num2 As Integer = 0
            Dim num3 As Integer = 1

            While num3 < array.Length
                Dim textChange1 As TextChange = array(num2)
                Dim textChange2 As TextChange = array(num3)
                Dim num4 As Integer = textChange2.Position - textChange1.OldEnd
                If num4 > 0 Then
                    textChange1.Position += num
                    num += textChange1.Delta
                    num2 = num3
                    num3 += 1
                    Continue While
                End If

                Catenate(textChange1.NewText, textChange2.NewText)
                If num4 = 0 Then
                    Catenate(textChange1.OldText, textChange2.OldText)
                ElseIf textChange1.OldEnd < textChange2.OldEnd Then
                    Dim startIndex As Integer = textChange1.OldEnd - textChange2.Position
                    Catenate(textChange1.OldText, textChange2.OldText.Substring(startIndex))
                End If
                array(num3) = Nothing
                num3 += 1
            End While

            array(num2).Position += num
            Dim list1 As New List(Of ITextChange)

            For i As Integer = 0 To array.Length - 1
                If array(i) IsNot Nothing Then
                    list1.Add(array(i))
                End If
            Next

            Return list1
        End Function
    End Class
End Namespace
