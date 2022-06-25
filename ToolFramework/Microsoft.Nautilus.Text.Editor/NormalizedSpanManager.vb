Imports System.Collections.Generic
Imports System.Windows.Media.TextFormatting
Imports Microsoft.Nautilus.Text.Classification

Namespace Microsoft.Nautilus.Text.Editor
    Friend Class NormalizedSpanManager
        Private Const _gapClass As String = "_gap_"
        Private _textSnapshot As ITextSnapshot
        Private _span As Span
        Private _classificationSpanList As IList(Of ClassificationSpan)
        Private _spaceNegotiationList As IList(Of SpaceNegotiation)
        Private _classificationFormatMap As IClassificationFormatMap
        Private _textClassificationType As IClassificationType
        Private _startNode As NormalizedSpan
        Private _containsBiDi As Boolean
        Private _predictedNode As NormalizedSpan
        Private Shared _lineBreak As New TextEndOfLine(1)

        Public ReadOnly Property VirtualCharacterPositions As List(Of Integer)

        Public ReadOnly Property ContainsBiDiCharacters As Boolean
            Get
                Return _containsBiDi
            End Get
        End Property

        Public Sub New(textSnapshot As ITextSnapshot, span As Span, classificationSpans As IList(Of ClassificationSpan), spaceNegotiations As IList(Of SpaceNegotiation), classificationFormatMap As IClassificationFormatMap, classificationTypeRegistry As IClassificationTypeRegistry)
            _textSnapshot = textSnapshot
            _span = span
            _classificationSpanList = classificationSpans
            _spaceNegotiationList = spaceNegotiations
            _classificationFormatMap = classificationFormatMap
            _textClassificationType = classificationTypeRegistry.GetClassificationType("text")
            _virtualCharacterPositions = New List(Of Integer)

            CreateNormalizedSpans()
            InsertSpaceNegotiationSpans()
            AddTextModifiers()
            MergeSpans()

            Dim curSpan = _startNode
            Dim start As Integer = 0

            While curSpan IsNot Nothing
                curSpan.StartCharacterIndex = start
                start += curSpan.Length
                curSpan = curSpan.Next
            End While

            _predictedNode = _startNode
        End Sub

        Public Function GetTextRun(characterIndex As Integer) As TextRun
            Dim span = _predictedNode
            While True
                If characterIndex >= span.StartCharacterIndex + span.Length Then
                    span = span.Next
                Else
                    If characterIndex >= span.StartCharacterIndex Then Exit While
                    span = span.Previous
                End If

                If span Is Nothing Then Return _lineBreak
            End While

            Dim textRun = span.GetTextRun(characterIndex)
            If span.Next IsNot Nothing Then
                _predictedNode = span.Next
            End If
            Return textRun
        End Function

        Private Sub CreateNormalizedSpans()
            Dim textProperties = _classificationFormatMap.GetTextProperties(_textClassificationType)
            Dim start = _span.Start
            _startNode = New NormalizedSpan("", "_gap_", _span.Start, textProperties)
            Dim normalizedSpan = _startNode

            For Each cSpan In _classificationSpanList
                Dim curSpan As Span = cSpan.GetSpan(_textSnapshot)
                If curSpan.Start >= _span.End Then Exit For

                Dim delta = curSpan.Start - start
                If delta > 0 Then
                    Dim span2 As New NormalizedSpan(_textSnapshot.GetText(New Span(start, delta)), "_gap_", start, textProperties)
                    normalizedSpan = normalizedSpan.AddNode(span2)
                    start = curSpan.Start

                ElseIf delta < 0 Then
                    Dim num3 As Integer = curSpan.End - start
                    If num3 <= 0 Then Continue For
                    curSpan = New Span(start, num3)

                End If

                If curSpan.End > _span.End Then
                    curSpan = New Span(curSpan.Start, _span.End - curSpan.Start)
                End If

                If curSpan.Length > 0 Then
                    Dim properties = _classificationFormatMap.GetTextProperties(cSpan.ClassificationType)
                    Dim span3 As New NormalizedSpan(_textSnapshot.GetText(curSpan), cSpan.ClassificationType.Classification, curSpan.Start, properties)
                    normalizedSpan = normalizedSpan.AddNode(span3)
                    start = curSpan.End
                    If start = _span.End Then Exit For
                End If
            Next

            Dim length = _span.End - start
            If length > 0 Then
                Dim span4 As New NormalizedSpan(_textSnapshot.GetText(New Span(start, length)), "_gap_", start, textProperties)
                normalizedSpan = normalizedSpan.AddNode(span4)
            End If

            If _startNode.Next IsNot Nothing Then
                _startNode = _startNode.Next
                _startNode.Previous = Nothing
            End If
        End Sub

        Private Sub InsertSpaceNegotiationSpans()
            If _spaceNegotiationList Is Nothing Then Return

            Dim normalizedSpan1 = _startNode
            For Each spaceNegotiation1 As SpaceNegotiation In _spaceNegotiationList
                Dim textRun1 As TextRun = New TextEmbeddedSpace(spaceNegotiation1.Size)
                normalizedSpan1 = normalizedSpan1.AddNode(New NormalizedSpan(textRun1, spaceNegotiation1.TextPosition))
                _virtualCharacterPositions.Add(spaceNegotiation1.TextPosition - _span.Start)
            Next
        End Sub

        Private Sub AddTextModifiers()
            Dim normalizedSpan1 As NormalizedSpan = _startNode
            While normalizedSpan1 IsNot Nothing
                If normalizedSpan1.AddTextModifiers() Then
                    _containsBiDi = True
                    _virtualCharacterPositions.Add(normalizedSpan1.StartCharacterIndex - _span.Start)
                    _virtualCharacterPositions.Add(normalizedSpan1.StartCharacterIndex + normalizedSpan1.Length - _span.Start)
                    normalizedSpan1 = normalizedSpan1.Next
                End If
                normalizedSpan1 = normalizedSpan1.Next
            End While

            If _startNode.Previous IsNot Nothing Then
                _startNode = _startNode.Previous
            End If
        End Sub

        Private Sub MergeSpans()
            Dim normalizedSpan1 As NormalizedSpan = _startNode
            While normalizedSpan1 IsNot Nothing

                normalizedSpan1 = normalizedSpan1.TryMergeNextSpan()
            End While
        End Sub

    End Class
End Namespace
