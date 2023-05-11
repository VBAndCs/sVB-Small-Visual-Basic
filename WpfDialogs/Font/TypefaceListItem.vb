Imports System.Text
Imports System.Globalization

Friend Class TypefaceListItem
    Inherits TextBlock
    Implements IComparable

    Private _displayName As String
    Private _simulated As Boolean

    Public Sub New(Family As FontFamily, ByVal FontWeight As FontWeight)
        Me.New(New Typeface(Family, FontStyles.Normal, FontWeight, FontStretches.Normal))
        _displayName = FontWeight.ToString
        Me.Text = _displayName
        Me.ToolTip = _displayName
    End Sub

    Public Sub New(Family As FontFamily, ByVal FontStyle As FontStyle)
        Me.New(New Typeface(Family, FontStyle, FontWeights.Normal, FontStretches.Normal))
        _displayName = FontStyle.ToString
        Me.Text = _displayName
        Me.ToolTip = _displayName
    End Sub

    Public Sub New(Family As FontFamily, ByVal FontStretch As FontStretch)
        Me.New(New Typeface(Family, FontStyles.Normal, FontWeights.Normal, FontStretch))
        _displayName = FontStretch.ToString
        Me.Text = _displayName
        Me.ToolTip = _displayName
    End Sub

    Public Sub New(ByVal typeface As Typeface)
        _displayName = GetDisplayName(typeface)
        _simulated = typeface.IsBoldSimulated OrElse typeface.IsObliqueSimulated

        Me.FontFamily = typeface.FontFamily
        Me.FontWeight = typeface.Weight
        Me.FontStyle = typeface.Style
        Me.FontStretch = typeface.Stretch

        Dim itemLabel As String = _displayName

        If _simulated Then
            Dim formatString As String = My.Resources.ResourceManager.GetString("simulated", CultureInfo.CurrentUICulture)
            itemLabel = String.Format(formatString, itemLabel)
        End If

        Me.Text = itemLabel
        Me.ToolTip = itemLabel

        ' In the case of symbol font, apply the default message font to the text so it can be read.
        If FontFamilyListItem.IsSymbolFont(typeface.FontFamily) Then
            Dim range As New TextRange(Me.ContentStart, Me.ContentEnd)
            range.ApplyPropertyValue(TextBlock.FontFamilyProperty, SystemFonts.MessageFontFamily)
        End If
    End Sub

    Public Overrides Function ToString() As String
        Return _displayName
    End Function

    Public ReadOnly Property Typeface() As Typeface
        Get
            Return New Typeface(FontFamily, FontStyle, FontWeight, FontStretch)
        End Get
    End Property

    Private Function IComparable_CompareTo(ByVal obj As Object) As Integer Implements IComparable.CompareTo
        Dim item As TypefaceListItem = TryCast(obj, TypefaceListItem)
        If item Is Nothing Then
            Return -1
        End If

        ' Sort all simulated faces after all non-simulated faces.
        If _simulated <> item._simulated Then
            Return If(_simulated, 1, -1)
        End If

        ' If weight differs then sort based on weight (lightest first).
        Dim difference As Integer = FontWeight.ToOpenTypeWeight() - item.FontWeight.ToOpenTypeWeight()
        If difference <> 0 Then
            Return If(difference > 0, 1, -1)
        End If

        ' If style differs then sort based on style (Normal, Italic, then Oblique).
        Dim thisStyle As FontStyle = FontStyle
        Dim otherStyle As FontStyle = item.FontStyle

        If thisStyle <> otherStyle Then
            If thisStyle = FontStyles.Normal Then
                ' This item is normal style and should come first.
                Return -1
            ElseIf otherStyle = FontStyles.Normal Then
                ' The other item is normal style and should come first.
                Return 1
            Else
                ' Neither is normal so sort italic before oblique.
                Return If(thisStyle = FontStyles.Italic, -1, 1)
            End If
        End If

        ' If stretch differs then sort based on stretch (Normal first, then numerically).
        Dim thisStretch As FontStretch = FontStretch
        Dim otherStretch As FontStretch = item.FontStretch

        If thisStretch <> otherStretch Then
            If thisStretch = FontStretches.Normal Then
                ' This item is normal stretch and should come first.
                Return -1
            ElseIf otherStretch = FontStretches.Normal Then
                ' The other item is normal stretch and should come first.
                Return 1
            Else
                ' Neither is normal so sort numerically.
                Return If(thisStretch.ToOpenTypeStretch() < otherStretch.ToOpenTypeStretch(), -1, 0)
            End If
        End If

        ' They're the same.
        Return 0
    End Function

    Friend Shared Function GetDisplayName(ByVal typeface As Typeface) As String
        Return NameDictionaryHelper.GetDisplayName(typeface.FaceNames)
    End Function

    Shared Sub [Select](Cmb As ComboBox, typeface As Windows.Media.Typeface)
        Dim Stretches = From Item As TypefaceListItem In Cmb.Items
                                   Select Item.FontStretch
                                   Distinct

        Dim Stretch As FontStretch
        If Stretches.Count = 1 Then
            Stretch = Stretches(0)
        Else
            Stretch = typeface.Stretch
        End If

        Dim Result = From Item As TypefaceListItem In Cmb.Items
                              Where Item.FontWeight = typeface.Weight AndAlso
                                         Item.FontStyle = typeface.Style AndAlso
                                         Item.FontStretch = Stretch

        If Result.Count > 0 Then
            Cmb.SelectedItem = Result(0)

        ElseIf typeface.Style = FontStyles.Italic Then
            Result = From Item As TypefaceListItem In Cmb.Items
                          Where Item.FontWeight = typeface.Weight AndAlso
                                     Item.FontStyle = FontStyles.Oblique AndAlso
                                     Item.FontStretch = Stretch

            If Result.Count > 0 Then Cmb.SelectedItem = Result(0)

        ElseIf typeface.Style = FontStyles.Oblique Then
            Result = From Item As TypefaceListItem In Cmb.Items
                          Where Item.FontWeight = typeface.Weight AndAlso
                                     Item.FontStyle = FontStyles.Italic AndAlso
                                     Item.FontStretch = Stretch

            If Result.Count > 0 Then Cmb.SelectedItem = Result(0)

        End If

    End Sub

End Class

