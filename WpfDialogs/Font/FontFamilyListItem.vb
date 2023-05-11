Imports System.Globalization

Friend Class FontFamilyListItem
    Inherits TextBlock
    Implements IComparable

    Private _displayName As String

    Public Sub New(ByVal fontFamily As FontFamily)        
        Me.FontFamily = fontFamily
        _displayName = GetDisplayName(fontFamily)
        Me.Text = _displayName
        Me.ToolTip = _displayName


        ' In the case of symbol font, apply the default message font to the text so it can be read.
        If IsSymbolFont(fontFamily) Then
            Dim range As New TextRange(Me.ContentStart, Me.ContentEnd)
            range.ApplyPropertyValue(TextBlock.FontFamilyProperty, SystemFonts.MessageFontFamily)
        End If
    End Sub

    Public Overrides Function ToString() As String
        Return _displayName
    End Function

    Private Function IComparable_CompareTo(ByVal obj As Object) As Integer Implements IComparable.CompareTo
        Return String.Compare(_displayName, obj.ToString(), True, CultureInfo.CurrentUICulture)
    End Function

    Friend Shared Function IsSymbolFont(ByVal fontFamily As FontFamily) As Boolean
        For Each typeface As Typeface In fontFamily.GetTypefaces()
            Dim face As GlyphTypeface
            If typeface.TryGetGlyphTypeface(face) Then
                Return face.Symbol
            End If
        Next typeface
        Return False
    End Function

    Friend Shared Function GetDisplayName(ByVal family As FontFamily) As String
        'If family.Source = "DiagramTTBlindAll" Then Stop
        Try
            Return NameDictionaryHelper.GetDisplayName(family.FamilyNames)
        Catch ex As Exception
            Return family.Source
        End Try
    End Function

    Shared Function CanShowFont(FontFamily As Windows.Media.FontFamily) As Boolean
        If FontDialog.ShowSymbolFonts Then Return True
        Return Not IsSymbolFont(FontFamily)
    End Function

End Class

