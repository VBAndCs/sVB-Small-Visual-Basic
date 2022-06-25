Imports System.Globalization

Namespace Microsoft.Nautilus.Text.Operations
    Friend Class NaturalLanguageNavigator
        Implements ITextStructureNavigator

        Private Enum CharacterType
            AlphaNumeric
            WhiteSpace
            Symbols
        End Enum

        Private _textBuffer As ITextBuffer
        Private ReadOnly _whiteSpaceCharacters As Char() = New Char(5) {" "c, vbTab(0), vbCr(0), vbLf(0), vbVerticalTab(0), vbFormFeed(0)}
        Private ReadOnly _extendedAlphaNumericCharacters As Char() = New Char(0) {"_"c}

        Public ReadOnly Property ContentType As String Implements ITextStructureNavigator.ContentType
            Get
                Return _textBuffer.ContentType
            End Get
        End Property

        Friend Sub New(textBuffer As ITextBuffer)
            _textBuffer = textBuffer
        End Sub

        Public Function GetPositionOfStartOfLine(currentPosition As SnapshotPoint) As Integer Implements ITextStructureNavigator.GetPositionOfStartOfLine
            If currentPosition.Snapshot.TextBuffer IsNot _textBuffer Then
                Throw New ArgumentException("currentPosition TextBuffer does not match with current TextBuffer")
            End If

            Dim containingLine As ITextSnapshotLine = currentPosition.GetContainingLine()
            Dim position1 As Integer = currentPosition.Position
            Dim start1 As Integer = containingLine.Start
            Dim positionOfNextNonWhiteSpaceCharacter As Integer = containingLine.GetPositionOfNextNonWhiteSpaceCharacter(0)
            Dim num As Integer = start1 + positionOfNextNonWhiteSpaceCharacter

            If position1 <> num AndAlso positionOfNextNonWhiteSpaceCharacter <> containingLine.Length Then
                Return num
            End If

            Return start1
        End Function

        Public Function GetPositionOfEndOfLine(currentPosition As SnapshotPoint) As Integer Implements ITextStructureNavigator.GetPositionOfEndOfLine
            If currentPosition.Snapshot.TextBuffer IsNot _textBuffer Then
                Throw New ArgumentException("currentPosition TextBuffer does not match with current TextBuffer")
            End If

            Return currentPosition.GetContainingLine().End
        End Function

        Public Function GetExtentOfWord(currentPosition As SnapshotPoint) As TextExtent Implements ITextStructureNavigator.GetExtentOfWord
            If currentPosition.Snapshot.TextBuffer IsNot _textBuffer Then
                Throw New ArgumentException("currentPosition  TextBuffer does not match with current TextBuffer")
            End If

            Dim position1 As Integer = currentPosition.Position
            If position1 = currentPosition.Snapshot.Length Then
                Return New TextExtent(position1, 0, isSignificant:=False)
            End If

            Dim containingLine As ITextSnapshotLine = currentPosition.GetContainingLine()
            Dim textIncludingLineBreak As String = containingLine.GetTextIncludingLineBreak()
            Dim characterType1 As CharacterType = GetCharacterType(currentPosition.GetChar())
            Dim num As Integer = FindStartOfWord(position1, containingLine, textIncludingLineBreak, characterType1)
            Dim num2 As Integer = FindEndOfWord(currentPosition.Snapshot, position1, containingLine, textIncludingLineBreak, characterType1)
            Return New TextExtent(num, num2 - num, characterType1 <> CharacterType.WhiteSpace)
        End Function

        Private Function GetCharacterType(c As Char) As CharacterType
            If Char.IsLetterOrDigit(c) Then
                Return CharacterType.AlphaNumeric
            End If

            If DoesArrayContainCharacter(_extendedAlphaNumericCharacters, c) Then
                Return CharacterType.AlphaNumeric
            End If

            If DoesArrayContainCharacter(_whiteSpaceCharacters, c) Then
                Return CharacterType.WhiteSpace
            End If

            If Char.IsSurrogate(c) OrElse UnicodeCategory.NonSpacingMark = Char.GetUnicodeCategory(c) Then
                Return CharacterType.AlphaNumeric
            End If

            Return CharacterType.Symbols
        End Function

        Private Shared Function DoesArrayContainCharacter(charArray As Char(), c As Char) As Boolean
            For i As Integer = 0 To charArray.Length - 1
                If charArray(i) = c Then
                    Return True
                End If
            Next

            Return False
        End Function

        Private Function GetPositionOfLastNonWhiteSpaceCharacter(text As String) As Integer
            Dim length1 As Integer = text.Length
            For num As Integer = length1 - 1 To 0 Step -1
                Dim c As Char = text(num)
                If GetCharacterType(c) <> CharacterType.WhiteSpace Then
                    Return num
                End If
            Next

            Return -1
        End Function

        Private Function FindEndOfWord(snapshot1 As ITextSnapshot, currentPosition As Integer, snapshotLine As ITextSnapshotLine, bufferLineText As String, baseType As CharacterType) As Integer
            If currentPosition >= snapshot1.Length - 1 Then
                Return snapshot1.Length
            End If

            Dim start1 As Integer = snapshotLine.Start
            Dim num As Integer = currentPosition

            While Threading.Interlocked.Increment(num) - start1 < bufferLineText.Length
                If GetCharacterType(bufferLineText(num - start1)) <> baseType Then
                    Return num
                End If
            End While

            If num = snapshot1.Length Then
                Return num
            End If

            Dim lineFromPosition As ITextSnapshotLine = snapshot1.GetLineFromPosition(snapshotLine.EndIncludingLineBreak)
            Dim positionOfNextNonWhiteSpaceCharacter As Integer = lineFromPosition.GetPositionOfNextNonWhiteSpaceCharacter(0)
            If positionOfNextNonWhiteSpaceCharacter <> lineFromPosition.LengthIncludingLineBreak Then
                Return num + positionOfNextNonWhiteSpaceCharacter
            End If

            Return num
        End Function

        Private Function FindStartOfWord(currentPosition As Integer, snapshotLine As ITextSnapshotLine, bufferLineText As String, baseType As CharacterType) As Integer
            If currentPosition = 0 Then Return 0

            Dim start1 As Integer = snapshotLine.Start
            Dim num As Integer = currentPosition

            While Threading.Interlocked.Decrement(num) > start1
                If GetCharacterType(bufferLineText(num - start1)) <> baseType Then
                    Return num + 1
                End If
            End While

            Return start1
        End Function
    End Class
End Namespace
