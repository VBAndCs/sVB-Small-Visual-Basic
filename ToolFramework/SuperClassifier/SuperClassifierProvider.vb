Imports System.Collections.Generic
Imports System.ComponentModel.Composition
Imports Microsoft.Nautilus.Core
Imports Microsoft.Nautilus.Text
Imports Microsoft.Nautilus.Text.Classification

Namespace SuperClassifier
    Public NotInheritable Class SuperClassifierProvider
        Private Shared _languageDataMap As Dictionary(Of String, LanguageData)
        Private _languageSpecifications As IEnumerable(Of LanguageSpecification)

        <Import>
        Public Property ClassificationTypeRegistry As IClassificationTypeRegistry

        <Import("LanguageSpecification")>
        Public Property LanguageSpecifications As IEnumerable(Of LanguageSpecification)
            Get
                Return Me._languageSpecifications
            End Get

            Set(value As IEnumerable(Of LanguageSpecification))
                _languageSpecifications = value
                _languageDataMap = Nothing
                BuildLanguageDataMap()
            End Set
        End Property

        <ContentType("text")>
        <Export(GetType(ClassifierProvider))>
        Public Function CreateClassifier(textBuffer As ITextBuffer) As IClassifier
            If _languageDataMap Is Nothing Then
                BuildLanguageDataMap()
            End If

            Dim tokenizer1 As Tokenizer = GetTokenizer(textBuffer)
            If tokenizer1 IsNot Nothing Then
                Return New SuperColorizer(textBuffer, ClassificationTypeRegistry, tokenizer1)
            End If

            Return Nothing
        End Function

        Friend Shared Function GetTokenizer(textBuffer As ITextBuffer) As Tokenizer
            If _languageDataMap Is Nothing Then Return Nothing
            Return New Tokenizer(GetLanguageData(textBuffer.ContentType), textBuffer)
        End Function

        Friend Shared Function GetLanguageData(contentType As String) As LanguageData
            If _languageDataMap Is Nothing Then Return Nothing

            Dim languageData As LanguageData = Nothing
            While Not _languageDataMap.TryGetValue(contentType, languageData)
                contentType = ContentTypeHelper.GetParentContentType(contentType)
                If contentType = "" Then Return Nothing
            End While

            Return languageData
        End Function

        Private Sub BuildLanguageDataMap()
            _languageDataMap = New Dictionary(Of String, LanguageData)
            For Each languageSpecification1 As LanguageSpecification In LanguageSpecifications
                If Not _languageDataMap.ContainsKey(languageSpecification1.ContentType) Then
                    _languageDataMap(languageSpecification1.ContentType) = New LanguageData(languageSpecification1)
                End If
            Next
        End Sub
    End Class
End Namespace
