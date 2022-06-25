Imports System.Collections.Generic

Namespace SuperClassifier
    Public Class LanguageSpecification

        Public Property ContentType As String

        Public ReadOnly Property Delimiters As IList(Of Delimiter) = New List(Of Delimiter)

        Public ReadOnly Property Literals As IList(Of Literals) = New List(Of Literals)

        Public ReadOnly Property Identifiers As IList(Of Identifiers) = New List(Of Identifiers)

        Public Property Indentors As Indentors
    End Class
End Namespace
