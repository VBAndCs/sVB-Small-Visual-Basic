Imports System.Collections.Generic
Imports System.ComponentModel.Composition
Imports System.Text
Imports Microsoft.Nautilus.Text.Classification.DataExports

Namespace Microsoft.Nautilus.Text.Classification
    <Export(GetType(IClassificationTypeRegistry))>
    Public NotInheritable Class ClassificationTypeRegistryImpl
        Implements IClassificationTypeRegistry

        Private _classificationTypes As Dictionary(Of String, ClassificationTypeImpl)
        Private _transientClassificationTypes As Dictionary(Of String, ClassificationTypeImpl)
        Private _classificationTypeDefinitions As IEnumerable(Of ClassificationTypeDefinition)

        <Import("ClassificationTypeDefinition")>
        Public Property ClassificationTypeDefinitions As IEnumerable(Of ClassificationTypeDefinition)
            Get
                Return _classificationTypeDefinitions
            End Get

            Set(value As IEnumerable(Of ClassificationTypeDefinition))
                _classificationTypeDefinitions = value
                _classificationTypes = Nothing
                _transientClassificationTypes = Nothing
                BuildClassificationTypes()
            End Set
        End Property

        Private ReadOnly Property TransientClassificationType As IClassificationType
            Get
                Return ClassificationTypes("(transient)")
            End Get
        End Property

        Private ReadOnly Property ClassificationTypes As Dictionary(Of String, ClassificationTypeImpl)
            Get
                If _classificationTypes Is Nothing Then
                    _classificationTypes = New Dictionary(Of String, ClassificationTypeImpl)
                End If

                Return _classificationTypes
            End Get
        End Property

        Public Function GetClassificationType(type As String) As IClassificationType Implements IClassificationTypeRegistry.GetClassificationType
            Dim value As ClassificationTypeImpl = Nothing
            BuildClassificationTypes()
            ClassificationTypes.TryGetValue(type, value)
            Return value
        End Function

        Public Function CreateTransientClassificationType(baseTypes As IEnumerable(Of IClassificationType)) As IClassificationType Implements IClassificationTypeRegistry.CreateTransientClassificationType
            If baseTypes Is Nothing Then
                Throw New ArgumentNullException("addionalTypes")
            End If

            If Not baseTypes.GetEnumerator().MoveNext() Then
                Throw New InvalidOperationException("Should not pass an empty IEnumerable<IClassificationType>")
            End If

            Return BuildTransientClassificationType(baseTypes)
        End Function

        Public Function CreateTransientClassificationType(ParamArray baseTypes As IClassificationType()) As IClassificationType Implements IClassificationTypeRegistry.CreateTransientClassificationType
            If baseTypes Is Nothing Then
                Throw New ArgumentNullException("addionalTypes")
            End If

            If baseTypes.Length = 0 Then
                Throw New InvalidOperationException("A transient classification type needs at least one base type.")
            End If

            Return BuildTransientClassificationType(baseTypes)
        End Function

        Private Sub BuildClassificationTypes()
            If _classificationTypes IsNot Nothing Then Return

            For Each definition In ClassificationTypeDefinitions
                Dim imp As ClassificationTypeImpl = Nothing
                If ClassificationTypes.ContainsKey(definition.Name) Then
                    imp = ClassificationTypes(definition.Name)
                Else
                    imp = New ClassificationTypeImpl(definition.Name)
                    ClassificationTypes.Add(definition.Name, imp)
                End If

                Dim derivesFrom1 As String = definition.DerivesFrom
                If derivesFrom1 IsNot Nothing Then
                    Dim classificationTypeImpl2 As ClassificationTypeImpl = Nothing
                    If ClassificationTypes.ContainsKey(derivesFrom1) Then
                        classificationTypeImpl2 = ClassificationTypes(derivesFrom1)
                    Else
                        classificationTypeImpl2 = New ClassificationTypeImpl(derivesFrom1)
                        ClassificationTypes.Add(derivesFrom1, classificationTypeImpl2)
                    End If
                    imp.AddBaseType(classificationTypeImpl2)
                End If
            Next
        End Sub

        Private Function BuildTransientClassificationType(baseTypes As IEnumerable(Of IClassificationType)) As IClassificationType
            If _transientClassificationTypes Is Nothing Then
                _transientClassificationTypes = New Dictionary(Of String, ClassificationTypeImpl)
            End If

            Dim sortedList1 As New SortedList(Of String, String)(StringComparer.Ordinal)
            For Each baseType As IClassificationType In baseTypes
                sortedList1.Add(baseType.Classification, baseType.Classification)
            Next

            Dim stringBuilder1 As New StringBuilder
            For Each key As String In sortedList1.Keys
                stringBuilder1.Append(key)
                stringBuilder1.Append(" - ")
            Next

            stringBuilder1.Append(TransientClassificationType.Classification)
            Dim value As ClassificationTypeImpl = Nothing
            If _transientClassificationTypes.TryGetValue(stringBuilder1.ToString(), value) Then
                Return value
            End If

            Dim classificationTypeImpl1 As New ClassificationTypeImpl(stringBuilder1.ToString())
            For Each baseType2 As IClassificationType In baseTypes
                classificationTypeImpl1.AddBaseType(baseType2)
            Next

            classificationTypeImpl1.AddBaseType(TransientClassificationType)
            _transientClassificationTypes(classificationTypeImpl1.Classification) = classificationTypeImpl1
            Return classificationTypeImpl1
        End Function

    End Class
End Namespace
