Imports System.Collections.Generic
Imports System.Globalization
Imports System.IO
Imports System.Text.RegularExpressions

Namespace Microsoft.Nautilus.Text.StringRebuilder
    Friend Class SimpleStringRebuilder
        Inherits BaseStringRebuilder
        Implements ITreeNode

        Friend NotInheritable Class LineBreak
            Friend Shared Function LengthOfLineBreak(c1 As Char, c2 As Char) As Integer
                Select Case c1
                    Case vbCr(0)
                        If c2 <> vbLf Then Return 1
                        Return 2

                    Case vbLf(0), vbVerticalTab(0), vbFormFeed(0), ChrW(&H85)
                        Return 1

                    Case Else
                        Return 0

                End Select

            End Function

            Friend Shared Function LengthOfLineBreak(c1 As Char) As Integer
                If c1 = vbCr OrElse c1 = vbLf OrElse c1 = vbVerticalTab OrElse c1 = vbFormFeed OrElse c1 = ChrW(&H85) Then
                    Return 1
                End If
                Return 0
            End Function
        End Class

        Private _span As Span
        Private _source As String
        Private _lineBreakSpans As IList(Of Span) = New List(Of Span)
        Private Shared _empty As IStringRebuilder = New SimpleStringRebuilder(Span.FromBounds(0, 0), String.Empty)
        Private Shared lonelyCrOrLfRegex As New Regex("((?<!" & vbCr & ")" & vbLf & ")|(" & vbCr & "(?!" & vbLf & "))", RegexOptions.Compiled)

        Public Overrides ReadOnly Property Length As Integer
            Get
                Return _span.Length
            End Get
        End Property

        Public Overrides ReadOnly Property LineBreakCount As Integer
            Get
                Return _lineBreakSpans.Count
            End Get
        End Property

        Default Public Overrides ReadOnly Property Item(index As Integer) As Char
            Get
                If index < 0 OrElse index >= Length Then
                    Throw New ArgumentOutOfRangeException("index")
                End If

                Return _source(index + _span.Start)
            End Get
        End Property

        Public ReadOnly Property Depth As Integer = 0 Implements ITreeNode.Depth

        Public ReadOnly Property StringRebuilder As IStringRebuilder Implements ITreeNode.StringRebuilder
            Get
                Return Me
            End Get
        End Property

        Friend Sub New(span As Span, source As String)
            _span = span
            _source = source
            Dim num As Integer = _span.Start

            While num < _span.End
                Dim num2 As Integer = (If((num + 1 < _span.End), LineBreak.LengthOfLineBreak(_source(num), _source(num + 1)), LineBreak.LengthOfLineBreak(_source(num))))
                If num2 = 0 Then
                    num += 1
                    Continue While
                End If

                _lineBreakSpans.Add(New Span(num - _span.Start, num2))
                num += num2
            End While
        End Sub

        Friend Sub New(span1 As Span, simpleSource As SimpleStringRebuilder)
            _span = New Span(span1.Start + simpleSource._span.Start, span1.Length)
            _source = simpleSource._source

            For i As Integer = simpleSource.GetLineNumberFromPosition(span1.Start) To simpleSource.LineBreakCount - 1
                Dim num As Integer = simpleSource._lineBreakSpans(i).Start - span1.Start
                If num >= span1.Length Then Exit For

                Dim [end] As Integer = Math.Min(span1.Length, simpleSource._lineBreakSpans(i).End - span1.Start)
                _lineBreakSpans.Add(Span.FromBounds(Math.Max(0, num), [end]))
            Next
        End Sub

        Friend Sub New(left As IStringRebuilder, right As IStringRebuilder)
            _span = New Span(0, left.Length + right.Length)
            _source = left.GetText(New Span(0, left.Length)) + right.GetText(New Span(0, right.Length))
            Dim num As Integer = 0

            If _source(left.Length) = vbLf AndAlso _source(left.Length - 1) = vbCr Then
                num = 1
            End If

            Dim num2 As Integer = left.LineBreakCount - num
            For i As Integer = 0 To num2 - 1
                Dim lineFromLineNumber As LineSpan = left.GetLineFromLineNumber(i)
                _lineBreakSpans.Add(New Span(lineFromLineNumber.End, lineFromLineNumber.LineBreakLength))
            Next

            If num = 1 Then
                _lineBreakSpans.Add(New Span(left.Length - 1, 2))
            End If

            For j As Integer = num To right.LineBreakCount - 1
                Dim lineFromLineNumber2 As LineSpan = right.GetLineFromLineNumber(j)
                _lineBreakSpans.Add(New Span(lineFromLineNumber2.End + left.Length, lineFromLineNumber2.LineBreakLength))
            Next
        End Sub

        Public Shared Function Create(text As String) As IStringRebuilder
            If text Is Nothing Then
                Throw New ArgumentNullException("text")
            End If

            If text.Length <> 0 Then
                Return New SimpleStringRebuilder(Span.FromBounds(0, text.Length), text)
            End If

            Return _empty
        End Function

        Public Shared Function Create(span1 As Span, text As String) As IStringRebuilder
            If text Is Nothing Then
                Throw New ArgumentNullException("text")
            End If

            If span1.End > text.Length Then
                Throw New ArgumentOutOfRangeException("span")
            End If

            If span1.Length <> 0 Then
                Return New SimpleStringRebuilder(span1, text)
            End If

            Return _empty
        End Function

        Public Overrides Function ToString() As String
            Return _source.Substring(_span.Start, _span.Length)
        End Function

        Public Overrides Function GetLineNumberFromPosition(position As Integer) As Integer
            If position < 0 OrElse position > Length Then
                Throw New ArgumentOutOfRangeException("position")
            End If

            Dim num As Integer = 0
            Dim num2 As Integer = _lineBreakSpans.Count

            While num < num2
                Dim num3 As Integer = (num + num2) \ 2
                If position < _lineBreakSpans(num3).End Then
                    num2 = num3
                Else
                    num = num3 + 1
                End If
            End While
            Return num
        End Function

        Public Overrides Function GetLineFromLineNumber(lineNumber As Integer) As LineSpan
            If lineNumber < 0 OrElse lineNumber > LineBreakCount Then
                Throw New ArgumentOutOfRangeException("lineNumber")
            End If

            Dim start1 As Integer = (If((lineNumber <> 0), _lineBreakSpans(lineNumber - 1).End, 0))
            Dim [end] As Integer
            Dim lineBreakLength1 As Integer

            If lineNumber < LineBreakCount Then
                [end] = _lineBreakSpans(lineNumber).Start
                lineBreakLength1 = _lineBreakSpans(lineNumber).Length
            Else
                [end] = _span.Length
                lineBreakLength1 = 0
            End If

            Return New LineSpan(lineNumber, Span.FromBounds(start1, [end]), lineBreakLength1)
        End Function

        Public Overrides Function CompareTo(other As IStringRebuilder, ignoreCase As Boolean, culture As CultureInfo) As Integer
            If culture Is Nothing Then
                Throw New ArgumentNullException("culture")
            End If

            If other Is Nothing Then
                Throw New ArgumentNullException("other")
            End If

            If Object.ReferenceEquals(Me, other) Then
                Return 0
            End If

            If Length < other.Length Then
                Dim num As Integer = CompareTo(other.Substring(Span.FromBounds(0, Length)), ignoreCase, culture)
                If num = 0 Then Return -1
                Return num
            End If

            If Length > other.Length Then
                Dim num2 As Integer = Substring(Span.FromBounds(0, other.Length)).CompareTo(other, ignoreCase, culture)
                If num2 = 0 Then Return 1
                Return num2
            End If

            If TypeOf other Is SimpleStringRebuilder Then
                Dim sb = CType(other, SimpleStringRebuilder)
                Return String.Compare(_source, _span.Start, sb._source, sb._span.Start, _span.Length, ignoreCase, culture)
            End If

            If TypeOf other Is BinaryStringRebuilder Then
                Return -other.CompareTo(Me, ignoreCase, culture)
            End If

            Return String.Compare(_source.Substring(_span.Start, _span.Length), other.GetText(New Span(0, other.Length)), ignoreCase, culture)
        End Function

        Public Overrides Function GetText(span1 As Span) As String
            If span1.End > Length Then
                Throw New ArgumentOutOfRangeException("span")
            End If

            Return _source.Substring(span1.Start + _span.Start, span1.Length)
        End Function

        Public Overrides Sub CopyTo(sourceIndex As Integer, destination As Char(), destinationIndex As Integer, count1 As Integer)
            If sourceIndex < 0 Then
                Throw New ArgumentOutOfRangeException("sourceIndex")
            End If

            If destination Is Nothing Then
                Throw New ArgumentNullException("destination")
            End If

            If destinationIndex < 0 Then
                Throw New ArgumentOutOfRangeException("destinationIndex")
            End If

            If count1 < 0 Then
                Throw New ArgumentOutOfRangeException("count")
            End If

            If sourceIndex + count1 > Length OrElse sourceIndex + count1 < 0 Then
                Throw New ArgumentOutOfRangeException("count")
            End If

            If destinationIndex + count1 > destination.Length OrElse destinationIndex + count1 < 0 Then
                Throw New ArgumentOutOfRangeException("count")
            End If

            _source.CopyTo(sourceIndex + _span.Start, destination, destinationIndex, count1)
        End Sub

        Public Overrides Sub Write(writer As TextWriter, span1 As Span)
            If writer Is Nothing Then
                Throw New ArgumentNullException("writer")
            End If

            If span1.End > Length Then
                Throw New ArgumentOutOfRangeException("span")
            End If

            Dim input As String = _source.Substring(span1.Start + _span.Start, span1.Length)
            writer.Write(lonelyCrOrLfRegex.Replace(input, Microsoft.VisualBasic.vbNewLine))
        End Sub

        Public Overrides Function Substring(span1 As Span) As IStringRebuilder
            If span1.End > Length Then
                Throw New ArgumentOutOfRangeException("span")
            End If

            If span1.Length = Length Then Return Me

            If span1.Length = 0 Then Return _empty

            Return New SimpleStringRebuilder(span1, Me)
        End Function

        Public Function Child(side1 As Side) As ITreeNode Implements ITreeNode.Child
            Throw New InvalidOperationException("No children.")
        End Function

        Public Function OtherChild(side1 As Side) As ITreeNode Implements ITreeNode.OtherChild
            Throw New InvalidOperationException("No children.")
        End Function
    End Class
End Namespace
