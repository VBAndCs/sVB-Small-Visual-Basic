Imports System.Collections.Generic
Imports System.Globalization
Imports System.Windows.Media.TextFormatting
Imports Microsoft.Nautilus.Text.Classification

Namespace Microsoft.Nautilus.Text.Editor
    Friend Class TextFormattingSource
        Inherits TextSource

        Private _normalizedSpanManager As NormalizedSpanManager

        Public ReadOnly Property VirtualCharacterPositions As IList(Of Integer)
            Get
                Return _normalizedSpanManager.VirtualCharacterPositions
            End Get
        End Property

        Public ReadOnly Property LineContainsBiDi As Boolean
            Get
                Return _normalizedSpanManager.ContainsBiDiCharacters
            End Get
        End Property

        Public Sub New(textSnapshot As ITextSnapshot, span As Span, classificationFormatMap As IClassificationFormatMap, classificationSpans As IList(Of ClassificationSpan), spaceNegotiations As IList(Of SpaceNegotiation), classificationTypeRegistry As IClassificationTypeRegistry)
            Dim list1 As List(Of SpaceNegotiation) = Nothing
            If spaceNegotiations IsNot Nothing Then
                list1 = New List(Of SpaceNegotiation)(spaceNegotiations)
                list1.Sort(New Comparison(Of SpaceNegotiation)(AddressOf Comparison))
            End If

            _normalizedSpanManager = New NormalizedSpanManager(textSnapshot, span, classificationSpans, list1, classificationFormatMap, classificationTypeRegistry)
        End Sub

        Private Function Comparison(x As SpaceNegotiation, y As SpaceNegotiation) As Integer
            Return x.TextPosition.CompareTo(y.TextPosition)
        End Function

        Public Overrides Function GetTextRun(textSourceCharacterIndex As Integer) As TextRun
            Return _normalizedSpanManager.GetTextRun(textSourceCharacterIndex)
        End Function

        Public Overrides Function GetPrecedingText(textSourceCharacterIndexLimit As Integer) As TextSpan(Of CultureSpecificCharacterBufferRange)
            Return New TextSpan(Of CultureSpecificCharacterBufferRange)(0, New CultureSpecificCharacterBufferRange(CultureInfo.CurrentUICulture, New CharacterBufferRange("", 0, 0)))
        End Function

        Public Overrides Function GetTextEffectCharacterIndexFromTextSourceCharacterIndex(textSourceCharacterIndex As Integer) As Integer
            Return 0
        End Function

        Public Shared Function ContainsBiDiCharacters(text As String) As Boolean
            For Each c As Char In text
                If (c >= ChrW(&H590) AndAlso c <= ChrW(&H8FF)) OrElse (c >= ChrW(&H200F) AndAlso c <= ChrW(&H202E)) OrElse c >= "יִ"c Then
                    Return True
                End If
            Next
            Return False
        End Function

        Public Shared Function IsBidiCharacter(c As Char) As Boolean
            If c < ChrW(&H590) Then Return False

            If (c >= ChrW(&H590) AndAlso c <= ChrW(&H8FF)) OrElse (c >= ChrW(&H200F) AndAlso c <= ChrW(&H202E)) OrElse c >= "יִ"c Then
                Return True
            End If

            Return False
        End Function
    End Class
End Namespace
