Imports System.Globalization
Imports System.Windows.Media.TextFormatting

Namespace Microsoft.Nautilus.Text.Editor
    Friend Class AvalonLineNumberMarginTextSource
        Inherits TextSource

        Private _text As String
        Private _textFormattingRunProperties As TextFormattingRunProperties

        Public Sub New(text As String, properties As TextFormattingRunProperties)
            If text Is Nothing Then
                Throw New ArgumentNullException("text")
            End If

            If properties Is Nothing Then
                Throw New ArgumentNullException("properties")
            End If

            _text = text
            _textFormattingRunProperties = properties
        End Sub

        Public Overrides Function GetPrecedingText(textSourceCharacterIndexLimit As Integer) As TextSpan(Of CultureSpecificCharacterBufferRange)
            Return New TextSpan(Of CultureSpecificCharacterBufferRange)(0, New CultureSpecificCharacterBufferRange(CultureInfo.CurrentUICulture, New CharacterBufferRange("", 0, 0)))
        End Function

        Public Overrides Function GetTextEffectCharacterIndexFromTextSourceCharacterIndex(textSourceCharacterIndex As Integer) As Integer
            If textSourceCharacterIndex < 0 OrElse textSourceCharacterIndex >= _text.Length Then
                Throw New ArgumentOutOfRangeException("textSourceCharacterIndex")
            End If

            Return textSourceCharacterIndex
        End Function

        Public Overrides Function GetTextRun(textSourceCharacterIndex As Integer) As TextRun
            If textSourceCharacterIndex < 0 OrElse textSourceCharacterIndex > _text.Length Then
                Throw New ArgumentOutOfRangeException("textSourceCharacterIndex")
            End If

            If textSourceCharacterIndex = _text.Length Then
                Return New TextEndOfLine(1)
            End If

            Return New TextCharacters(_text.Substring(textSourceCharacterIndex), _textFormattingRunProperties)
        End Function
    End Class
End Namespace
