Imports System.Windows
Imports System.Windows.Media
Imports System.Windows.Media.TextFormatting

Namespace Microsoft.Nautilus.Text.Editor
    Public Class TextLine
        Implements IDisposable

        Public Shared LineSpacing As Double = 4.0

        Dim _textLine As System.Windows.Media.TextFormatting.TextLine

        Public Sub New(textLine As System.Windows.Media.TextFormatting.TextLine)
            _textLine = textLine
        End Sub

        Public ReadOnly Property Width As Double
            Get
                Return _textLine.Width
            End Get
        End Property

        Public ReadOnly Property Height As Double
            Get
                Return _textLine.Height + LineSpacing
            End Get
        End Property

        Public ReadOnly Property WidthIncludingTrailingWhitespace As Double
            Get
                Return _textLine.WidthIncludingTrailingWhitespace
            End Get
        End Property

        Public ReadOnly Property OverhangTrailing As Double
            Get
                Return _textLine.OverhangTrailing
            End Get
        End Property

        Public ReadOnly Property OverhangLeading As Double
            Get
                Return _textLine.OverhangLeading
            End Get
        End Property

        Public ReadOnly Property OverhangAfter As Double
            Get
                Return _textLine.OverhangAfter
            End Get
        End Property

        Public ReadOnly Property Extent As Double
            Get
                Return _textLine.Extent
            End Get
        End Property

        Public ReadOnly Property Length As Integer
            Get
                Return _textLine.Length
            End Get
        End Property

        Public ReadOnly Property NewlineLength As Integer
            Get
                Return _textLine.NewlineLength
            End Get
        End Property

        Public ReadOnly Property Baseline As Double
            Get
                Return _textLine.Baseline
            End Get
        End Property

        Friend Sub Draw(drawingContext As DrawingContext, origin As Point, inversion As InvertAxes)
            _textLine.Draw(drawingContext, origin, inversion)
        End Sub

        Friend Function GetIndexedGlyphRuns() As IEnumerable(Of IndexedGlyphRun)
            Return _textLine.GetIndexedGlyphRuns()
        End Function

        Friend Function GetTextBounds(firstTextSourceCharacterIndex As Integer, textLength As Integer) As IList(Of TextFormatting.TextBounds)
            Return _textLine.GetTextBounds(firstTextSourceCharacterIndex, textLength)
        End Function

        Friend Function GetDistanceFromCharacterHit(characterHit As CharacterHit) As Double
            Return _textLine.GetDistanceFromCharacterHit(characterHit)
        End Function

        Friend Function GetPreviousCaretCharacterHit(characterHit As CharacterHit) As CharacterHit
            Return _textLine.GetPreviousCaretCharacterHit(characterHit)
        End Function

        Friend Function GetNextCaretCharacterHit(characterHit As CharacterHit) As CharacterHit
            Return _textLine.GetNextCaretCharacterHit(characterHit)
        End Function

        Friend Function GetCharacterHitFromDistance(distance As Double) As CharacterHit
            Return _textLine.GetCharacterHitFromDistance(distance)
        End Function

        Public Sub Dispose() Implements IDisposable.Dispose
            _textLine.Dispose()
        End Sub
    End Class
End Namespace
