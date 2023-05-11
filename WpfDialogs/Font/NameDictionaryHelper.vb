Imports System.Globalization
Imports System.Windows.Markup

Friend NotInheritable Class NameDictionaryHelper

    Private Sub New()
    End Sub

    Public Shared Function GetDisplayName(ByVal nameDictionary As LanguageSpecificStringDictionary) As String
        ' Look up the display name based on the UI culture, which is the same culture
        ' used for resource loading.
        Dim userLanguage As XmlLanguage = XmlLanguage.GetLanguage(CultureInfo.CurrentUICulture.IetfLanguageTag)

        ' Look for an exact match.
        Dim name As String
        If nameDictionary.TryGetValue(userLanguage, name) Then
            Return name
        End If

        ' No exact match; return the name for the most closely related language.
        Dim bestRelatedness As Integer = -1
        Dim bestName As String = String.Empty

        For Each pair As KeyValuePair(Of XmlLanguage, String) In nameDictionary
            Dim relatedness As Integer = GetRelatedness(pair.Key, userLanguage)
            If relatedness > bestRelatedness Then
                bestRelatedness = relatedness
                bestName = pair.Value
            End If
        Next pair

        Return bestName
    End Function

    Public Shared Function GetDisplayName(ByVal nameDictionary As IDictionary(Of CultureInfo, String)) As String
        ' Look for an exact match.
        Dim name As String
        If nameDictionary.TryGetValue(CultureInfo.CurrentUICulture, name) Then
            Return name
        End If

        ' No exact match; return the name for the most closely related language.
        Dim bestRelatedness As Integer = -1
        Dim bestName As String = String.Empty

        Dim userLanguage As XmlLanguage = XmlLanguage.GetLanguage(CultureInfo.CurrentUICulture.IetfLanguageTag)

        For Each pair As KeyValuePair(Of CultureInfo, String) In nameDictionary
            Dim relatedness As Integer = GetRelatedness(XmlLanguage.GetLanguage(pair.Key.IetfLanguageTag), userLanguage)
            If relatedness > bestRelatedness Then
                bestRelatedness = relatedness
                bestName = pair.Value
            End If
        Next pair

        Return bestName
    End Function

    Private Shared Function GetRelatedness(ByVal keyLang As XmlLanguage, ByVal userLang As XmlLanguage) As Integer
        Try
            ' Get equivalent cultures.
            Dim keyCulture As CultureInfo = CultureInfo.GetCultureInfoByIetfLanguageTag(keyLang.IetfLanguageTag)
            Dim userCulture As CultureInfo = CultureInfo.GetCultureInfoByIetfLanguageTag(userLang.IetfLanguageTag)
            If Not userCulture.IsNeutralCulture Then
                userCulture = userCulture.Parent
            End If

            ' If the key is a prefix or parent of the user language it's a good match.
            If IsPrefixOf(keyLang.IetfLanguageTag, userLang.IetfLanguageTag) OrElse userCulture.Equals(keyCulture) Then
                Return 2
            End If

            ' If the key and user language share a common prefix or parent neutral culture, it's a reasonable match.
            If IsPrefixOf(TrimSuffix(userLang.IetfLanguageTag), keyLang.IetfLanguageTag) OrElse userCulture.Equals(keyCulture.Parent) Then
                Return 1
            End If
        Catch e1 As ArgumentException
            ' Language tag with no corresponding CultureInfo.
        End Try

        ' They're unrelated languages.
        Return 0
    End Function

    Private Shared Function TrimSuffix(ByVal tag As String) As String
        Dim i As Integer = tag.LastIndexOf("-"c)
        If i > 0 Then
            Return tag.Substring(0, i)
        Else
            Return tag
        End If
    End Function

    Private Shared Function IsPrefixOf(ByVal prefix As String, ByVal tag As String) As Boolean
        Return prefix.Length < tag.Length AndAlso tag.Chars(prefix.Length) = "-"c AndAlso String.CompareOrdinal(prefix, 0, tag, 0, prefix.Length) = 0
    End Function
End Class

