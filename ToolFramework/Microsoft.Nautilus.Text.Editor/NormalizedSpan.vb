Imports System.Windows.Media.TextFormatting

Namespace Microsoft.Nautilus.Text.Editor
    Friend Class NormalizedSpan
        Private _classification As String
        Private _text As String
        Private _properties As TextFormattingRunProperties
        Private _textRun As TextRun
        Private _canSplitOrMerge As Boolean

        Public Property [Next] As NormalizedSpan

        Public Property Previous As NormalizedSpan

        Public ReadOnly Property Length As Integer

        Public Property StartCharacterIndex As Integer

        Public Sub New(text As String, classification As String, startCharacterIndex1 As Integer, properties As TextFormattingRunProperties)
            _text = text
            _Length = text.Length
            _classification = classification
            _StartCharacterIndex = startCharacterIndex1
            _properties = properties
            _Previous = Nothing
            _Next = Nothing
            _canSplitOrMerge = True
        End Sub

        Public Sub New(textRun As TextRun, startCharacterIndex As Integer)
            _textRun = textRun
            _Length = _textRun.Length
            _StartCharacterIndex = startCharacterIndex
        End Sub

        Public Function GetTextRun(startCharacterIndex As Integer) As TextRun
            If _textRun IsNot Nothing Then
                Return _textRun
            End If

            Dim characterString As String = _text
            If startCharacterIndex > _StartCharacterIndex Then
                characterString = _text.Substring(startCharacterIndex - _StartCharacterIndex)
            End If

            Return New TextCharacters(characterString, _properties)
        End Function

        Public Function AddTextModifiers() As Boolean
            If _canSplitOrMerge AndAlso TextFormattingSource.ContainsBiDiCharacters(_text) Then
                Dim normalizedSpan1 As New NormalizedSpan(New TextFormattingModifier(_properties), StartCharacterIndex)
                normalizedSpan1.Next = Me
                normalizedSpan1.Previous = Previous

                If Previous IsNot Nothing Then
                    Previous.Next = normalizedSpan1
                End If

                Previous = normalizedSpan1
                Dim normalizedSpan2 As New NormalizedSpan(New TextEndOfSegment(1), StartCharacterIndex + Length)
                normalizedSpan2.Next = _Next
                normalizedSpan2.Previous = Me

                If _Next IsNot Nothing Then
                    _Next.Previous = normalizedSpan2
                End If

                _Next = normalizedSpan2
                Return True
            End If

            Return False
        End Function

        Public Function AddNode(span As NormalizedSpan) As NormalizedSpan
            If span.StartCharacterIndex > StartCharacterIndex + Length Then
                Return _Next.AddNode(span)
            End If

            If _canSplitOrMerge AndAlso span.StartCharacterIndex < StartCharacterIndex + Length Then
                Dim startCharacterIndex1 As Integer = span.StartCharacterIndex
                Dim text As String = _text.Substring(startCharacterIndex1 - StartCharacterIndex)
                Dim span2 As New NormalizedSpan(text, _classification, startCharacterIndex1, _properties)

                _Length = startCharacterIndex1 - StartCharacterIndex
                _text = _text.Substring(0, _Length)
                AddNode(span)
                Return span.AddNode(span2)
            End If

            span.Previous = Me
            If _Next IsNot Nothing Then
                span.Next = _Next
                _Next.Previous = span
            End If

            _Next = span
            Return span
        End Function

        Public Function TryMergeNextSpan() As NormalizedSpan
            If Not _canSplitOrMerge OrElse _Next Is Nothing OrElse Not _Next._canSplitOrMerge Then
                Return _Next
            End If

            Dim flag As Boolean = False
            If _classification = _Next._classification OrElse _properties Is _Next._properties Then
                flag = True

            ElseIf _properties.SameSize(_Next._properties) Then
                If _classification = "whitespace" Then
                    flag = True
                    _classification = _Next._classification
                    _properties = _Next._properties
                ElseIf _Next._classification = "whitespace" Then
                    flag = True
                End If
            End If

            If flag Then
                _text &= _Next._text
                _Length += _Next._Length
                _Next = _Next.Next

                If _Next IsNot Nothing Then
                    _Next.Previous = Me
                End If

                Return Me
            End If

            Return _Next
        End Function

    End Class
End Namespace
