Imports System.Collections.Generic

Namespace SuperClassifier

    Friend NotInheritable Class LanguageData

        Private Shared ReadOnly emptyCollection As IEnumerable(Of String) = New String() {}
        Friend ReadOnly ContentType As String
        Friend LexExpressions As List(Of LexExpression)
        Friend IdentifierLexExpression As New IdentifierLexExpression(Nothing, Nothing, isCaseSensitive:=False)

        Friend ReadOnly Property IndentingWords As IEnumerable(Of String) = emptyCollection

        Friend ReadOnly Property UnindentingWords As IEnumerable(Of String) = emptyCollection

        Friend ReadOnly Property IsCaseSensitive As Boolean
            Get
                Return IdentifierLexExpression.IsCaseSensitive
            End Get
        End Property

        Friend Sub New(languageSpecification As LanguageSpecification)
            ContentType = languageSpecification.ContentType
            LexExpressions = New List(Of LexExpression)

            For Each delimiter As Delimiter In languageSpecification.Delimiters
                AddDelimitedLexExpression(delimiter.Start, delimiter.End, delimiter.IsMultiLine, delimiter.Ignore, delimiter.Class)
            Next

            For Each literal As Literals In languageSpecification.Literals
                AddLiteralExpression(literal.LiteralList, literal.Class)
            Next

            For Each identifier As Identifiers In languageSpecification.Identifiers
                AddIdentifierExpression(identifier.PrefixCharacters, identifier.BodyCharacters, identifier.IsCaseSensitive)
                AddKeywords(identifier.Keywords)
            Next

            If languageSpecification.Indentors IsNot Nothing Then
                AddIndentors(languageSpecification.Indentors.Indenting, languageSpecification.Indentors.Unindenting)
            End If
        End Sub

        Friend Sub AddIdentifierExpression(prefixChars As String, bodyChars As String, isCaseSensitive As Boolean)
            IdentifierLexExpression = New IdentifierLexExpression(prefixChars, bodyChars, isCaseSensitive)
        End Sub

        Friend Sub AddKeywords(keywords1 As IList(Of String))
            IdentifierLexExpression.AddKeywords(keywords1)
        End Sub

        Friend Sub AddLiteralExpression(literals1 As IList(Of String), classification As String)
            LexExpressions.Add(New LiteralLexExpression(literals1, classification))
        End Sub

        Friend Sub AddDelimitedLexExpression(start1 As String, [end] As String, isMultiLine1 As Boolean, ignore1 As String, classification As String)
            LexExpressions.Add(New DelimitedLexExpression(start1, [end], isMultiLine1, ignore1, classification))
        End Sub

        Friend Sub AddIndentors(indentingWords As IList(Of String), unindentingWords As IList(Of String))
            _IndentingWords = indentingWords
            _UnindentingWords = unindentingWords
        End Sub
    End Class
End Namespace
