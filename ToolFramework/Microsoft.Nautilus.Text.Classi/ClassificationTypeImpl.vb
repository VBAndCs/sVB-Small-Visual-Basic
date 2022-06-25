Imports System.Collections.Generic

Namespace Microsoft.Nautilus.Text.Classification
    Friend Class ClassificationTypeImpl
        Implements IClassificationType

        Private _name As String
        Private BaseTypesList As New List(Of IClassificationType)

        Public ReadOnly Property Classification As String Implements IClassificationType.Classification
            Get
                Return _name
            End Get
        End Property

        Public ReadOnly Property BaseTypes As IEnumerable(Of IClassificationType) Implements IClassificationType.BaseTypes
            Get
                Return BaseTypesList.AsReadOnly()
            End Get
        End Property

        Friend Sub New(name As String)
            _name = name
        End Sub

        Friend Sub AddBaseType(baseType As IClassificationType)
            BaseTypesList.Add(baseType)
        End Sub

        Public Function IsOfType(type As String) As Boolean Implements IClassificationType.IsOfType
            If _name = type Then
                Return True
            End If

            For Each baseTypes As IClassificationType In BaseTypesList
                If baseTypes.IsOfType(type) Then Return True
            Next

            Return False
        End Function

        Public Overrides Function ToString() As String
            Return _name
        End Function

    End Class
End Namespace
