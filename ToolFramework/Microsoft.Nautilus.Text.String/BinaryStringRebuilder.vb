Imports System.Globalization
Imports System.IO

Namespace Microsoft.Nautilus.Text.StringRebuilder
    Friend Class BinaryStringRebuilder
        Inherits BaseStringRebuilder
        Implements ITreeNode

        Friend Class SimpleTreeNode
            Implements ITreeNode
            Private _source As IStringRebuilder

            Public ReadOnly Property Depth As Integer = 0 Implements ITreeNode.Depth

            Public ReadOnly Property StringRebuilder As IStringRebuilder Implements ITreeNode.StringRebuilder
                Get
                    Return _source
                End Get
            End Property

            Public Sub New(source As IStringRebuilder)
                If source Is Nothing Then
                    Throw New ArgumentNullException("source")
                End If

                _source = source
            End Sub

            Public Overrides Function ToString() As String
                Return _source.ToString()
            End Function

            Public Function Child(side1 As Side) As ITreeNode Implements ITreeNode.Child
                Throw New InvalidOperationException("No children.")
            End Function

            Public Function OtherChild(side1 As Side) As ITreeNode Implements ITreeNode.OtherChild
                Throw New InvalidOperationException("No children.")
            End Function
        End Class

        Private _left As ITreeNode
        Private _right As ITreeNode

        Private Shared _crlf As IStringRebuilder = SimpleStringRebuilder.Create(Microsoft.VisualBasic.vbNewLine)
        Friend Shared _maxCharactersToConsolidate As Integer = 200
        Friend Shared _maxLinesToConsolidate As Integer = 8

        Public Overrides ReadOnly Property Length As Integer

        Public Overrides ReadOnly Property LineBreakCount As Integer

        Default Public Overrides ReadOnly Property item(index As Integer) As Char
            Get
                If index < 0 OrElse index >= Length Then
                    Throw New ArgumentOutOfRangeException("index")
                End If

                If index >= _left.StringRebuilder.Length Then
                    Return _right.StringRebuilder(index - _left.StringRebuilder.Length)
                End If

                Return _left.StringRebuilder(index)
            End Get
        End Property

        Public ReadOnly Property Depth As Integer Implements ITreeNode.Depth

        Public ReadOnly Property StringRebuilder As IStringRebuilder Implements ITreeNode.StringRebuilder
            Get
                Return Me
            End Get
        End Property

        Friend Sub New(left As IStringRebuilder, right As IStringRebuilder)
            _left = TryCast(left, ITreeNode)
            If _left Is Nothing Then
                _left = New SimpleTreeNode(left)
            End If

            _right = TryCast(right, ITreeNode)
            If _right Is Nothing Then
                _right = New SimpleTreeNode(right)
            End If

            _depth = 1 + Math.Max(_left.Depth, _right.Depth)
            _length = left.Length + right.Length
            _lineBreakCount = left.LineBreakCount + right.LineBreakCount
        End Sub

        Private Function Balance() As IStringRebuilder
            If _left.Depth > _right.Depth + 1 Then
                Return Pivot(Side.Left)
            End If

            If _right.Depth > _left.Depth + 1 Then
                Return Pivot(Side.Right)
            End If

            Return Me
        End Function

        Private Function Pivot(deepSide As Side) As IStringRebuilder
            Dim treeNode As ITreeNode = Child(deepSide)
            Dim treeNode2 As ITreeNode = OtherChild(deepSide)
            Dim treeNode3 As ITreeNode = treeNode.Child(deepSide)
            Dim treeNode4 As ITreeNode = treeNode.OtherChild(deepSide)

            If treeNode3.Depth >= treeNode4.Depth Then
                Dim left As IStringRebuilder
                If deepSide = Side.Right Then
                    left = Create(treeNode2.StringRebuilder, treeNode4.StringRebuilder)
                    Return Create(left, treeNode3.StringRebuilder)
                End If

                left = Create(treeNode4.StringRebuilder, treeNode2.StringRebuilder)
                Return Create(treeNode3.StringRebuilder, left)
            End If

            Dim treeNode5 As ITreeNode = treeNode4.Child(deepSide)
            Dim treeNode6 As ITreeNode = treeNode4.OtherChild(deepSide)
            Dim left2 As IStringRebuilder
            Dim right As IStringRebuilder

            If deepSide = Side.Right Then
                left2 = Create(treeNode2.StringRebuilder, treeNode6.StringRebuilder)
                right = Create(treeNode5.StringRebuilder, treeNode3.StringRebuilder)
                Return Create(left2, right)
            End If

            left2 = Create(treeNode6.StringRebuilder, treeNode2.StringRebuilder)
            right = Create(treeNode3.StringRebuilder, treeNode5.StringRebuilder)
            Return Create(right, left2)
        End Function

        Public Shared Function Create(left As IStringRebuilder, right As IStringRebuilder) As IStringRebuilder
            If left Is Nothing Then
                Throw New ArgumentNullException("left")
            End If

            If right Is Nothing Then
                Throw New ArgumentNullException("right")
            End If

            If left.Length = 0 Then Return right

            If right.Length = 0 Then Return left

            If left.Length + right.Length < _maxCharactersToConsolidate AndAlso left.LineBreakCount + right.LineBreakCount <= _maxLinesToConsolidate Then
                Return New SimpleStringRebuilder(left, right)
            End If

            If right(0) = vbLf AndAlso left(left.Length - 1) = vbCr Then
                Return Create(Create(left.Substring(Span.FromBounds(0, left.Length - 1)), _crlf), right.Substring(Span.FromBounds(1, right.Length)))
            End If

            Return New BinaryStringRebuilder(left, right).Balance()
        End Function

        Public Overrides Function ToString() As String
            Return String.Format(
                CultureInfo.InvariantCulture,
                 If((_depth Mod 2 = 0), "({0})({1})", "[{0}][{1}]"),
                 New Object(1) {
                            _left.StringRebuilder.ToString(),
                            _right.StringRebuilder.ToString()
                 })
        End Function

        Public Overrides Function GetLineNumberFromPosition(position As Integer) As Integer
            If position < 0 OrElse position > Length Then
                Throw New ArgumentOutOfRangeException("position")
            End If

            If position > _left.StringRebuilder.Length Then
                Return _left.StringRebuilder.LineBreakCount + _right.StringRebuilder.GetLineNumberFromPosition(position - _left.StringRebuilder.Length)
            End If

            Return _left.StringRebuilder.GetLineNumberFromPosition(position)
        End Function

        Public Overrides Function GetLineFromLineNumber(lineNumber As Integer) As LineSpan
            If lineNumber < 0 OrElse lineNumber > LineBreakCount Then
                Throw New ArgumentOutOfRangeException("lineNumber")
            End If

            If lineNumber < _left.StringRebuilder.LineBreakCount Then
                Return _left.StringRebuilder.GetLineFromLineNumber(lineNumber)
            End If

            If lineNumber > _left.StringRebuilder.LineBreakCount Then
                Dim lineFromLineNumber As LineSpan = _right.StringRebuilder.GetLineFromLineNumber(lineNumber - _left.StringRebuilder.LineBreakCount)
                Return New LineSpan(lineNumber, New Span(lineFromLineNumber.Start + _left.StringRebuilder.Length, lineFromLineNumber.Length), lineFromLineNumber.LineBreakLength)
            End If

            Dim start1 As Integer = (If((lineNumber <> 0), _left.StringRebuilder.GetLineFromLineNumber(lineNumber).Start, 0))
            Dim [end] As Integer
            Dim lineBreakLength1 As Integer

            If lineNumber = LineBreakCount Then
                [end] = Length
                lineBreakLength1 = 0
            Else
                Dim lineFromLineNumber2 As LineSpan = _right.StringRebuilder.GetLineFromLineNumber(0)
                [end] = lineFromLineNumber2.End + _left.StringRebuilder.Length
                lineBreakLength1 = lineFromLineNumber2.LineBreakLength
            End If

            Return New LineSpan(lineNumber, Span.FromBounds(start1, [end]), lineBreakLength1)
        End Function

        Public Overrides Function CompareTo(other As IStringRebuilder, ignoreCase As Boolean, culture As CultureInfo) As Integer
            If culture Is Nothing Then
                Throw New ArgumentNullException("culture")
            End If

            If Object.ReferenceEquals(Me, other) Then Return 0

            If other Is Nothing Then
                Throw New ArgumentNullException("other")
            End If

            Dim num As Integer
            If other.Length <= _left.StringRebuilder.Length Then
                num = _left.StringRebuilder.Substring(Span.FromBounds(0, other.Length)).CompareTo(other, ignoreCase, culture)
                If num = 0 AndAlso Length > other.Length Then
                    num = 1
                End If
            Else
                num = _left.StringRebuilder.CompareTo(other.Substring(Span.FromBounds(0, _left.StringRebuilder.Length)), ignoreCase, culture)
                If num = 0 Then
                    num = _right.StringRebuilder.CompareTo(other.Substring(Span.FromBounds(_left.StringRebuilder.Length, other.Length)), ignoreCase, culture)
                End If
            End If

            Return num
        End Function

        Public Overrides Function GetText(span As Span) As String
            If span.End > Length Then
                Throw New ArgumentOutOfRangeException("span")
            End If

            If span.End <= _left.StringRebuilder.Length Then
                Return _left.StringRebuilder.GetText(span)
            End If

            If span.Start >= _left.StringRebuilder.Length Then
                Return _right.StringRebuilder.GetText(New Span(span.Start - _left.StringRebuilder.Length, span.Length))
            End If

            Dim array As Char() = New Char(span.Length - 1) {}
            Dim num As Integer = _left.StringRebuilder.Length - span.Start
            _left.StringRebuilder.CopyTo(span.Start, array, 0, num)
            _right.StringRebuilder.CopyTo(0, array, num, span.Length - num)

            Return New String(array)
        End Function

        Public Overrides Sub CopyTo(sourceIndex As Integer, destination As Char(), destinationIndex As Integer, count As Integer)
            If sourceIndex >= _left.StringRebuilder.Length Then
                _right.StringRebuilder.CopyTo(sourceIndex - _left.StringRebuilder.Length, destination, destinationIndex, count)
                Return
            End If

            If sourceIndex + count <= _left.StringRebuilder.Length Then
                _left.StringRebuilder.CopyTo(sourceIndex, destination, destinationIndex, count)
                Return
            End If

            Dim num As Integer = _left.StringRebuilder.Length - sourceIndex
            _left.StringRebuilder.CopyTo(sourceIndex, destination, destinationIndex, num)
            _right.StringRebuilder.CopyTo(0, destination, destinationIndex + num, count - num)
        End Sub

        Public Overrides Sub Write(writer As TextWriter, span1 As Span)
            If writer Is Nothing Then
                Throw New ArgumentNullException("writer")
            End If

            If span1.End > Length Then
                Throw New ArgumentOutOfRangeException("span")
            End If

            If span1.Start >= _left.StringRebuilder.Length Then
                _right.StringRebuilder.Write(writer, New Span(span1.Start - _left.StringRebuilder.Length, span1.Length))
                Return
            End If

            If span1.End <= _left.StringRebuilder.Length Then
                _left.StringRebuilder.Write(writer, span1)
                Return
            End If

            _left.StringRebuilder.Write(writer, Span.FromBounds(span1.Start, _left.StringRebuilder.Length))
            _right.StringRebuilder.Write(writer, Span.FromBounds(0, span1.End - _left.StringRebuilder.Length))
        End Sub

        Public Overrides Function Substring(span1 As Span) As IStringRebuilder
            If span1.End > Length Then
                Throw New ArgumentOutOfRangeException("span")
            End If

            If span1.Length = Length Then
                Return Me
            End If

            If span1.End <= _left.StringRebuilder.Length Then
                Return _left.StringRebuilder.Substring(span1)
            End If

            If span1.Start >= _left.StringRebuilder.Length Then
                Return _right.StringRebuilder.Substring(New Span(span1.Start - _left.StringRebuilder.Length, span1.Length))
            End If

            Return Create(_left.StringRebuilder.Substring(Span.FromBounds(span1.Start, _left.StringRebuilder.Length)), _right.StringRebuilder.Substring(Span.FromBounds(0, span1.End - _left.StringRebuilder.Length)))
        End Function

        Public Function Child(side1 As Side) As ITreeNode Implements ITreeNode.Child
            If side1 <> 0 Then
                Return _right
            End If

            Return _left
        End Function

        Public Function OtherChild(side1 As Side) As ITreeNode Implements ITreeNode.OtherChild
            If side1 <> 0 Then Return _left

            Return _right
        End Function
    End Class
End Namespace
