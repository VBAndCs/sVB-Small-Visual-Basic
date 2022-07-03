Imports System.Collections.Generic
Imports System.Collections.ObjectModel
Imports System.Text

Namespace Microsoft.Nautilus.Text
    Public NotInheritable Class NormalizedSpanCollection
        Inherits ReadOnlyCollection(Of Span)

        Private Class OrderedSpanList
            Inherits List(Of Span)
        End Class

        Public Sub New()
            MyBase.New(CType(New Span() {}, IList(Of Span)))
        End Sub

        Public Sub New(span As Span)
            MyBase.New(CType(New Span(0) {span}, IList(Of Span)))
        End Sub

        Public Sub New(spans As IEnumerable(Of Span))
            MyBase.New(NormalizeSpans(spans))
        End Sub

        Public Shared Function Union(set1 As NormalizedSpanCollection, set2 As NormalizedSpanCollection) As NormalizedSpanCollection
            If set1 Is Nothing Then
                Throw New ArgumentNullException("set1")
            End If

            If set2 Is Nothing Then
                Throw New ArgumentNullException("set2")
            End If

            If set1.Count = 0 Then
                Return set2
            End If

            If set2.Count = 0 Then
                Return set1
            End If

            Dim orderedSpanList1 As New OrderedSpanList
            Dim i As Integer = 0
            Dim j As Integer = 0
            Dim start As Integer = -1
            Dim [end] As Integer = Integer.MaxValue

            While i < set1.Count AndAlso j < set2.Count
                Dim span1 As Span = set1(i)
                Dim span2 As Span = set2(j)
                If span1.Start < span2.Start Then
                    UpdateSpanUnion(span1, orderedSpanList1, start, [end])
                    i += 1
                Else
                    UpdateSpanUnion(span2, orderedSpanList1, start, [end])
                    j += 1
                End If
            End While

            While i < set1.Count
                UpdateSpanUnion(set1(i), orderedSpanList1, start, [end])
                i += 1
            End While

            While j < set2.Count
                UpdateSpanUnion(set2(j), orderedSpanList1, start, [end])
                j += 1
            End While

            If [end] <> Integer.MaxValue Then
                orderedSpanList1.Add(New Span(start, [end] - start))
            End If

            Return New NormalizedSpanCollection(orderedSpanList1)
        End Function

        Public Shared Function Intersection(set1 As NormalizedSpanCollection, set2 As NormalizedSpanCollection) As NormalizedSpanCollection
            If set1 Is Nothing Then
                Throw New ArgumentNullException("set1")
            End If

            If set2 Is Nothing Then
                Throw New ArgumentNullException("set2")
            End If

            If set1.Count = 0 Then
                Return set1
            End If

            If set2.Count = 0 Then
                Return set2
            End If

            Dim orderedSpanList1 As New OrderedSpanList
            Dim num As Integer = 0
            Dim num2 As Integer = 0

            While num < set1.Count AndAlso num2 < set2.Count
                Dim span1 As Span = set1(num)
                Dim span2 As Span = set2(num2)
                Dim num3 As Integer = Math.Max(span1.Start, span2.Start)
                Dim num4 As Integer = Math.Min(span1.End, span2.End)

                If num3 < num4 Then
                    orderedSpanList1.Add(New Span(num3, num4 - num3))
                End If

                If span1.End < span2.End Then
                    num += 1
                Else
                    num2 += 1
                End If

            End While
            Return New NormalizedSpanCollection(orderedSpanList1)
        End Function

        Public Shared Operator =(set1 As NormalizedSpanCollection, set2 As NormalizedSpanCollection) As Boolean
            If Object.ReferenceEquals(set1, set2) Then
                Return True
            End If

            If Object.ReferenceEquals(set1, Nothing) OrElse Object.ReferenceEquals(set2, Nothing) Then
                Return False
            End If

            If set1.Count <> set2.Count Then
                Return False
            End If

            For i As Integer = 0 To set1.Count - 1
                If set1(i) <> set2(i) Then
                    Return False
                End If
            Next

            Return True
        End Operator

        Public Shared Operator <>(set1 As NormalizedSpanCollection, set2 As NormalizedSpanCollection) As Boolean
            Return Not (set1 Is set2)
        End Operator

        Public Function Intersects([set] As NormalizedSpanCollection) As Boolean
            If [set] Is Nothing Then
                Throw New ArgumentNullException("set")
            End If

            Dim num As Integer = 0
            Dim num2 As Integer = 0
            While num < MyBase.Count AndAlso num2 < [set].Count
                Dim span1 As Span = MyBase.Item(num)
                Dim span2 As Span = [set](num2)
                If span1.End > span2.Start AndAlso span2.End > span1.Start Then
                    Return True
                End If

                If span1.End < span2.End Then
                    num += 1
                Else
                    num2 += 1
                End If
            End While
            Return False
        End Function

        Public Overrides Function GetHashCode() As Integer
            Dim num As Integer = 0
            Using enumerator As IEnumerator(Of Span) = GetEnumerator()
            While enumerator.MoveNext()
                num = num Xor enumerator.Current.GetHashCode()
            End While
            Return num
            End Using
        End Function
        Public Overrides Function Equals(obj As Object) As Boolean
            Dim normalizedSpanCollection1 As NormalizedSpanCollection = TryCast(obj, NormalizedSpanCollection)
            Return Me Is normalizedSpanCollection1
        End Function
        Public Overrides Function ToString() As String
            Dim stringBuilder1 As New StringBuilder("{")
            Using enumerator As IEnumerator(Of Span) = GetEnumerator()
                While enumerator.MoveNext()
                    stringBuilder1.Append(enumerator.Current.ToString())
                End While
            End Using
            stringBuilder1.Append("}")
            Return stringBuilder1.ToString()
        End Function

        Private Sub New(normalizedSpans As OrderedSpanList)
            MyBase.New(CType(normalizedSpans, IList(Of Span)))
        End Sub
        Private Shared Sub UpdateSpanUnion(span1 As Span, spans As IList(Of Span), ByRef start As Integer, ByRef [end] As Integer)
            If [end] < span1.Start Then
                spans.Add(New Span(start, [end] - start))
                start = -1
                [end] = Integer.MaxValue
            End If
            If [end] = Integer.MaxValue Then
                start = span1.Start
                [end] = span1.End
            Else
                [end] = Math.Max([end], span1.End)
            End If
        End Sub

        Private Shared Function NormalizeSpans(spans As IEnumerable(Of Span)) As IList(Of Span)
            If spans Is Nothing Then
                Throw New ArgumentNullException("spans")
            End If

            Dim spanList = spans.ToList()
            If spanList.Count <= 1 Then Return spanList

            spanList.Sort(Function(s1 As Span, s2 As Span) s1.Start.CompareTo(s2.Start))

            Dim result As New List(Of Span)(spanList.Count)
            Dim start1 = spanList(0).Start
            Dim end1 = spanList(0).End

            For i As Integer = 1 To spanList.Count - 1
                Dim start2 = spanList(i).Start
                Dim end2 = spanList(i).End

                If end1 < start2 Then
                    result.Add(New Span(start1, end1 - start1))
                    start1 = start2
                    end1 = end2
                Else
                    end1 = Math.Max(end1, end2)
                End If
            Next

            result.Add(New Span(start1, end1 - start1))
            Return result
        End Function
    End Class
End Namespace
