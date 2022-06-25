Imports System.Collections.Generic
Imports System.Collections.ObjectModel
Imports System.ComponentModel.Composition
Imports System.Globalization
Imports System.Windows
Imports System.Windows.Media
Imports Microsoft.Nautilus.Core
Imports Microsoft.Nautilus.Text.Classification.DataExports
Imports Microsoft.Nautilus.Text.Editor

Namespace Microsoft.Nautilus.Text.Classification

    <Export(GetType(IClassificationFormatMap))>
    Public NotInheritable Class ClassificationFormatMap
        Implements IClassificationFormatMap

        Private Class ClassificationTypeOrder
            Implements IOrderable

            Private _after As String
            Private _before As String

            Public ReadOnly Iterator Property After As IEnumerable(Of String) Implements IOrderable.After
                Get
                    If _after IsNot Nothing Then
                        Yield _after
                    End If
                End Get
            End Property

            Public ReadOnly Iterator Property Before As IEnumerable(Of String) Implements IOrderable.Before
                Get
                    If _before IsNot Nothing Then
                        Yield _before
                    End If
                End Get
            End Property

            Public ReadOnly Property Name As String Implements IOrderable.Name

            Public Sub New(name As String, before As String, after As String)
                _Name = name
                _before = before
                _after = after
            End Sub
        End Class

        Private Const _aboveDefaultPriority As String = "{above default priority}"
        Private _ambientProperties As AmbientClassificationFormatMap
        Private _formats As IEnumerable(Of ClassificationFormat)
        Private _priorityOrder As List(Of IClassificationType)
        Private _textFormattingPropertiesWithDefaults As New Dictionary(Of String, TextFormattingRunProperties)

        <Import>
        Public Property ClassificationTypeRegistry As IClassificationTypeRegistry

        Public ReadOnly Property CurrentPriorityOrder As ReadOnlyCollection(Of IClassificationType) Implements IClassificationFormatMap.CurrentPriorityOrder
            Get
                If _priorityOrder Is Nothing Then
                    _priorityOrder = New List(Of IClassificationType)
                    Dim list1 = CreateDefaultOrdering()
                    AddAllProvidedFormatsToOrdering(list1)
                    list1 = New List(Of ClassificationTypeOrder)(_Orderer.Order(list1))
                    CreateFormatPriorityOrder(list1)
                End If

                Return _priorityOrder.AsReadOnly()
            End Get
        End Property

        <Import("ClassificationFormat")>
        Public Property Formats As IEnumerable(Of ClassificationFormat)
            Get
                Return _formats
            End Get

            Set(value As IEnumerable(Of ClassificationFormat))
                _formats = value
                _textFormattingPropertiesWithDefaults.Clear()
                GetDefaultProperties()
            End Set
        End Property

        <Import>
        Public Property Orderer As IOrderer

        Public Sub New()
            _ambientProperties = New AmbientClassificationFormatMap
        End Sub

        Private Sub AddAllProvidedFormatsToOrdering(orders As IList(Of ClassificationTypeOrder))
            For Each format As ClassificationFormat In Formats
                Dim before1 As String
                Dim after1 As String
                If format.After IsNot Nothing AndAlso format.Before IsNot Nothing Then
                    before1 = format.Before
                    after1 = format.After
                Else
                    before1 = "{above default priority}"
                    after1 = "Default Priority"
                End If

                For Each classificationType As String In format.ClassificationTypes
                    AddToPriorityList(classificationType, after1, before1, orders)
                Next
            Next
        End Sub

        Private Shared Sub AddToPriorityList(name As String, after As String, before As String, orders As IList(Of ClassificationTypeOrder))
            orders.Add(New ClassificationTypeOrder(name, before, after))
        End Sub

        Private Shared Function CreateDefaultOrdering() As IList(Of ClassificationTypeOrder)
            Dim list1 As IList(Of ClassificationTypeOrder) = New List(Of ClassificationTypeOrder)
            AddToPriorityList("Low Priority", Nothing, "Default Priority", list1)
            AddToPriorityList("Default Priority", "Low Priority", "{above default priority}", list1)
            AddToPriorityList("{above default priority}", "Default Priority", "High Priority", list1)
            AddToPriorityList("High Priority", "{above default priority}", Nothing, list1)
            Return list1
        End Function

        Private Sub CreateFormatPriorityOrder(orders As IList(Of ClassificationTypeOrder))
            For Each order1 As ClassificationTypeOrder In orders
                If order1.Name <> "Default Priority" AndAlso order1.Name <> "High Priority" AndAlso order1.Name <> "Low Priority" AndAlso order1.Name <> "{above default priority}" Then
                    _priorityOrder.Add(_ClassificationTypeRegistry.GetClassificationType(order1.Name))
                End If
            Next
        End Sub

        Private Shared Function CreateTextPropertiesFromProvision(provision As ClassificationFormat) As TextFormattingRunProperties
            Dim backgroundBrush1 As Brush = provision.BackgroundBrush
            Dim foregroundBrush1 As Brush = provision.ForegroundBrush
            Dim typeface1 As Typeface = CreateTypeFace(provision.FontFamily, provision.FontStyle, provision.FontWeight, provision.FontStretch, provision.FallbackFontFamily)
            Dim size As Double? = Nothing

            If provision.FontRenderingSize.HasValue Then
                size = provision.FontRenderingSize.Value
            End If

            Dim hintingSize As Double? = Nothing
            If provision.FontHintingSize.HasValue Then
                hintingSize = provision.FontHintingSize.Value
            End If

            Dim textDecorations1 = provision.TextDecorations
            Dim textEffects1 = provision.TextEffects
            Dim cultureInfo1 As CultureInfo = Nothing
            Return TextFormattingRunProperties.CreateTextFormattingRunProperties(cultureInfo:=If((Not provision.Culture.HasValue), CultureInfo.CurrentCulture, New CultureInfo(provision.Culture.Value)), foreground:=foregroundBrush1, background:=backgroundBrush1, typeface:=typeface1, size:=size, hintingSize:=hintingSize, textDecorations:=textDecorations1, textEffects:=textEffects1)
        End Function

        Private Shared Function CreateTypeFace(fontFamily1 As String, fontStyle1 As FontStyle, fontWeight1 As FontWeight, fontStretch1 As FontStretch, fallbackFontFamily1 As String) As Typeface
            If fontFamily1 Is Nothing Then
                Return Nothing
            End If

            Dim fontFamily2 As New FontFamily(fontFamily1)
            If fallbackFontFamily1 Is Nothing Then
                Return New Typeface(fontFamily2, fontStyle1, fontWeight1, fontStretch1)
            End If

            Return New Typeface(fontFamily2, fontStyle1, fontWeight1, fontStretch1, New FontFamily(fallbackFontFamily1))
        End Function

        Private Function FindProvision(classification As String) As ClassificationFormat
            For Each format As ClassificationFormat In Formats
                For Each classificationType As String In format.ClassificationTypes
                    If classificationType = classification Then
                        Return format
                    End If
                Next
            Next

            Return Nothing
        End Function

        Private Sub GetDefaultProperties()
            Dim value As TextFormattingRunProperties = Nothing
            If Not _ambientProperties.Properties.TryGetValue("text", value) Then
                value = GetTextPropertiesFromProviders("text")
            End If
            If value Is Nothing Then
                value = TextFormattingRunProperties.CreateTextFormattingRunProperties(New Typeface(SystemFonts.MenuFontFamily, SystemFonts.MenuFontStyle, SystemFonts.MenuFontWeight, FontStretches.Normal), SystemFonts.MenuFontSize, SystemColors.MenuTextColor)
            End If
            _textFormattingPropertiesWithDefaults("text") = value
        End Sub

        Public Function GetTextProperties(classificationType As IClassificationType) As TextFormattingRunProperties Implements IClassificationFormatMap.GetTextProperties
            If classificationType Is Nothing Then
                Throw New ArgumentNullException("classificationType")
            End If

            Dim value As TextFormattingRunProperties = Nothing
            If _textFormattingPropertiesWithDefaults.TryGetValue(classificationType.Classification, value) Then
                Return value
            End If

            value = TextFormattingRunProperties.MergeProperties(GetTextPropertiesFromHierarchy(classificationType), _textFormattingPropertiesWithDefaults("text"))
            _textFormattingPropertiesWithDefaults(classificationType.Classification) = value
            Return value
        End Function

        Private Function GetTextPropertiesFromHierarchy(classificationType As IClassificationType) As TextFormattingRunProperties
            Dim value As TextFormattingRunProperties = Nothing
            If _ambientProperties.Properties.TryGetValue(classificationType.Classification, value) Then
                Return value
            End If

            value = GetTextPropertiesFromProviders(classificationType.Classification)
            Dim list1 As New List(Of IClassificationType)(classificationType.BaseTypes)
            Dim count1 As Integer = list1.Count
            If count1 = 1 Then
                value = TextFormattingRunProperties.MergeProperties(value, GetTextPropertiesFromHierarchy(list1(0)))
            ElseIf count1 > 1 Then
                Dim num As Integer = 0
                For num2 As Integer = CurrentPriorityOrder.Count - 1 To 0 Step -1
                    If list1.Contains(CurrentPriorityOrder(num2)) Then
                        value = TextFormattingRunProperties.MergeProperties(value, GetTextPropertiesFromHierarchy(CurrentPriorityOrder(num2)))
                        num += 1
                        If count1 = num Then Exit For
                    End If
                Next
            End If
            Return value
        End Function

        Private Function GetTextPropertiesFromProviders(classification As String) As TextFormattingRunProperties
            Dim classificationFormat1 As ClassificationFormat = FindProvision(classification)
            Dim result As TextFormattingRunProperties = Nothing
            If classificationFormat1 IsNot Nothing Then
                result = CreateTextPropertiesFromProvision(classificationFormat1)
            End If
            Return result
        End Function

        Public Sub SetTextProperties(classificationType As IClassificationType, properties1 As TextFormattingRunProperties) Implements IClassificationFormatMap.SetTextProperties
            If classificationType Is Nothing Then
                Throw New ArgumentNullException("classificationType")
            End If
            If properties1 Is Nothing Then
                Throw New ArgumentNullException("properties")
            End If
            _ambientProperties.Properties(classificationType.Classification) = properties1
            _ambientProperties.Update()
            _textFormattingPropertiesWithDefaults.Clear()
            GetDefaultProperties()
        End Sub

        Public Sub SwapPriorities(firstType As IClassificationType, secondType As IClassificationType) Implements IClassificationFormatMap.SwapPriorities
            Dim index As Integer = _priorityOrder.IndexOf(firstType)
            Dim index2 As Integer = _priorityOrder.IndexOf(secondType)
            _priorityOrder(index) = secondType
            _priorityOrder(index2) = firstType
            _textFormattingPropertiesWithDefaults.Clear()
            GetDefaultProperties()
        End Sub

    End Class
End Namespace
