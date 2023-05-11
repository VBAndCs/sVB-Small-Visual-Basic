' Based on https://github.com/microsoft/wpf-samples/tree/main/Sample%20Applications/FontDialog

Imports System.Globalization
Imports System.Windows.Threading
Imports System.Windows.Controls.Primitives

Public Class FontChooser

#Region "Private fields and types"

    Private _familyCollection As ICollection(Of FontFamily) ' see FamilyCollection property
    Private _defaultSampleText As String = "Text sample  " & "نموذج"
    Private _previewSampleText As String
    Private _pointsText As String

    Private _updatePending As Boolean ' indicates a call to OnUpdate is scheduled
    Private _familyListValid As Boolean ' indicates the list of font families is valid
    Private _typefaceListValid As Boolean ' indicates the list of typefaces is valid
    Private _typefaceListSelectionValid As Boolean ' indicates the current selection in the typeface list is valid
    Private _previewValid As Boolean ' indicates the preview control is valid
    Private _tabDictionary As Dictionary(Of TabItem, TabState) ' state and logic for each tab
    Private _currentFeature As DependencyProperty
    Private _currentFeaturePage As TypographyFeaturePage

    Private Shared ReadOnly CommonlyUsedFontSizes() As Double = {3.0, 4.0, 5.0, 6.0, 7.0, 8.0, 9.0, 10.0, 11.0, 12.0, 12.5, 13.0, 13.5, 14.0, 15.0, 16.0, 17.0, 18.0, 19.0, 20.0, 22.0, 24.0, 26.0, 28.0, 30.0, 32.0, 34.0, 36.0, 38.0, 40.0, 42.0, 44.0, 46.0, 48.0, 50.0, 52.0, 56.0, 60.0, 64.0, 68.0, 72.0, 76.0, 80.0}

    ' Specialized metadata object for font chooser dependency properties
    Private Class FontPropertyMetadata
        Inherits FrameworkPropertyMetadata

        Public ReadOnly TargetProperty As DependencyProperty

        Public Sub New(ByVal defaultValue As Object, ByVal changeCallback As PropertyChangedCallback, ByVal targetProperty As DependencyProperty)
            MyBase.New(defaultValue, changeCallback)
            Me.TargetProperty = targetProperty
        End Sub
    End Class

    ' Specialized metadata object for typographic font chooser properties
    Private Class TypographicPropertyMetadata
        Inherits FontPropertyMetadata

        Public Sub New(ByVal defaultValue As Object, ByVal targetProperty As DependencyProperty, ByVal featurePage As TypographyFeaturePage, ByVal sampleTextTag As String)
            MyBase.New(defaultValue, _callback, targetProperty)
            Me.FeaturePage = featurePage
            Me.SampleTextTag = sampleTextTag
        End Sub

        Public ReadOnly FeaturePage As TypographyFeaturePage
        Public ReadOnly SampleTextTag As String

        Private Shared _callback As New PropertyChangedCallback(AddressOf FontChooser.TypographicPropertyChangedCallback)
    End Class

    ' Object used to initialize the right-hand side of the typographic properties tab
    Private Class TypographyFeaturePage
        Public Sub New(ByVal items() As Item)
            Me.Items = items
        End Sub

        Public Sub New(ByVal enumType As Type)
            Dim names() As String = System.Enum.GetNames(enumType)
            Dim values As Array = System.Enum.GetValues(enumType)

            Items = New Item(names.Length - 1) {}

            For i As Integer = 0 To names.Length - 1
                Items(i) = New Item(names(i), values.GetValue(i))
            Next i
        End Sub

        Public Shared ReadOnly BooleanFeaturePage As New TypographyFeaturePage(New Item() {New Item("Disabled", False), New Item("Enabled", True)})

        Public Shared ReadOnly IntegerFeaturePage As New TypographyFeaturePage(New Item() {New Item("_0", 0), New Item("_1", 1), New Item("_2", 2), New Item("_3", 3), New Item("_4", 4), New Item("_5", 5), New Item("_6", 6), New Item("_7", 7), New Item("_8", 8), New Item("_9", 9)})

        Public Structure Item
            Public Sub New(ByVal tag As String, ByVal value As Object)
                Me.Tag = tag
                Me.Value = value
            End Sub
            Public ReadOnly Tag As String
            Public ReadOnly Value As Object
        End Structure

        Public ReadOnly Items() As Item
    End Class

    Private Delegate Sub UpdateCallback()

    ' Encapsulates the state and initialization logic of a tab control item.
    Private Class TabState
        Public Sub New(ByVal initMethod As UpdateCallback)
            InitializeTab = initMethod
        End Sub

        Public IsValid As Boolean = False
        Public ReadOnly InitializeTab As UpdateCallback
    End Class

#End Region

#Region "Construction and initialization"

    Public Sub New()
        InitializeComponent()
    End Sub

    Protected Overrides Sub OnInitialized(ByVal e As EventArgs)
        MyBase.OnInitialized(e)
        _previewSampleText = _defaultSampleText
        _pointsText = typefaceNameRun.Text

        ' Initialize the dictionary that maps tab control items to handler objects.
        _tabDictionary = New Dictionary(Of TabItem, TabState)(tabControl.Items.Count)
        _tabDictionary.Add(samplesTab, New TabState(New UpdateCallback(AddressOf InitializeSamplesTab)))
        _tabDictionary.Add(typographyTab, New TabState(New UpdateCallback(AddressOf InitializeTypographyTab)))
        _tabDictionary.Add(descriptiveTextTab, New TabState(New UpdateCallback(AddressOf InitializeDescriptiveTextTab)))

        ' Hook up events for the tab control.
        AddHandler tabControl.SelectionChanged, AddressOf tabControl_SelectionChanged

        ' Initialize the list of font sizes and select the nearest size.
        For Each value As Double In CommonlyUsedFontSizes
            sizeList.Items.Add(value)
        Next value
        OnSelectedFontSizeChanged(FontSizeInPoints)

        previewTextBox.Text = _previewSampleText
        ' Schedule background updates.
        ScheduleUpdate()
    End Sub


#End Region

#Region "Event handlers"

    Private Sub OnOKButtonClicked(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Dim Wnd = Window.GetWindow(Me)
        Wnd.DialogResult = True
        Wnd.Close()
    End Sub

    Private Sub OnCancelButtonClicked(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Dim Wnd = Window.GetWindow(Me)
        Wnd.Close()
    End Sub

    Private _fontFamilyTextBoxSelectionStart As Integer

    Private Sub fontFamilyTextBox_SelectionChanged(ByVal sender As Object, ByVal e As RoutedEventArgs)
        _fontFamilyTextBoxSelectionStart = fontFamilyTextBox.SelectionStart
    End Sub

    Private Sub fontFamilyTextBox_TextChanged(ByVal sender As Object, ByVal e As TextChangedEventArgs)
        Dim text As String = fontFamilyList.Text
        fontFamilyList.IsDropDownOpen = True

        ScrollIntoView(fontFamilyList)
    End Sub

    Private Sub ComboxTextBox_PreviewKeyDown(ByVal sender As Object, ByVal e As KeyEventArgs)
        OnComboBoxPreviewKeyDown(sender, e)
    End Sub

    Private Sub fontFamilyList_SelectionChanged(ByVal sender As Object, ByVal e As SelectionChangedEventArgs)
        Dim item As FontFamilyListItem = TryCast(fontFamilyList.SelectedItem, FontFamilyListItem)
        If item IsNot Nothing Then
            SelectedFontFamily = item.FontFamily
        End If
    End Sub

    Private Sub sizeList_SelectionChanged(ByVal sender As Object, ByVal e As SelectionChangedEventArgs)
        If sizeList.SelectedIndex = -1 Then Return
        FontSizeInPoints = sizeList.SelectedItem
    End Sub

    Dim ExitTypefaceListSelectionChanged As Boolean = False

    Private Sub typefaceList_SelectionChanged(ByVal sender As Object, ByVal e As SelectionChangedEventArgs)
        If ExitTypefaceListSelectionChanged Then Return

        Dim item As TypefaceListItem = TryCast(typefaceList.SelectedItem, TypefaceListItem)
        If item IsNot Nothing Then
            SelectedFontWeight = item.FontWeight
            SelectedFontStyle = item.FontStyle
            SelectedFontStretch = item.FontStretch
        End If
    End Sub

    Private Sub textDecorationCheckStateChanged(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Dim Decorations As New TextDecorationCollection()

        If underlineCheckBox.IsChecked.Value Then
            Decorations.Add(System.Windows.TextDecorations.Underline(0))
        End If
        If baselineCheckBox.IsChecked.Value Then
            Decorations.Add(System.Windows.TextDecorations.Baseline(0))
        End If
        If strikethroughCheckBox.IsChecked.Value Then
            Decorations.Add(TextDecorations.Strikethrough(0))
        End If
        If overlineCheckBox.IsChecked.Value Then
            Decorations.Add(TextDecorations.OverLine(0))
        End If

        Decorations.Freeze()
        SelectedTextDecorations = Decorations
    End Sub

    Private Sub tabControl_SelectionChanged(ByVal sender As Object, ByVal e As SelectionChangedEventArgs)
        Dim tab As TabState = CurrentTabState
        If tab IsNot Nothing AndAlso (Not tab.IsValid) Then
            tab.InitializeTab()
            tab.IsValid = True
        End If
    End Sub

    Private Sub featureList_SelectionChanged(ByVal sender As Object, ByVal e As SelectionChangedEventArgs)
        InitializeTypographyTab()
    End Sub

#End Region

#Region "Public properties and methods"

    ''' <summary>
    ''' Collection of font families to display in the font family list. By default this is Fonts.SystemFontFamilies,
    ''' but a client could set this to another collection returned by Fonts.GetFontFamilies, e.g., a collection of
    ''' application-defined fonts.
    ''' </summary>
    Public Property FontFamilyCollection() As ICollection(Of FontFamily)
        Get
            Return If(_familyCollection Is Nothing, Fonts.SystemFontFamilies, _familyCollection)
        End Get

        Set(ByVal value As ICollection(Of FontFamily))
            If value IsNot _familyCollection Then
                _familyCollection = value
                InvalidateFontFamilyList()
            End If
        End Set
    End Property

    Friend Shared Function GetFontProperties() As List(Of DependencyProperty)
        Dim Props As New List(Of DependencyProperty)
        For Each [property] As DependencyProperty In FontProperties
            Dim metadata As FontPropertyMetadata = TryCast([property].GetMetadata(GetType(FontChooser)), FontPropertyMetadata)
            If metadata IsNot Nothing Then Props.Add(metadata.TargetProperty)
        Next
        Return Props
    End Function

    ''' <summary>
    ''' Sets the font chooser selection properties to match the properites of the specified object.
    ''' </summary>
    Public Sub SetPropertiesFromObject(ByVal obj As DependencyObject)
        For Each [property] As DependencyProperty In FontProperties
            Dim metadata As FontPropertyMetadata = TryCast([property].GetMetadata(GetType(FontChooser)), FontPropertyMetadata)
            If metadata IsNot Nothing Then
                Me.SetValue([property], obj.GetValue(metadata.TargetProperty))
            End If
        Next [property]
    End Sub

    ''' <summary>
    ''' Sets the properites of the specified object to match the font chooser selection properties.
    ''' </summary>
    Public Sub ApplyPropertiesToObject(ByVal obj As DependencyObject)
        For Each [property] As DependencyProperty In FontProperties
            Dim metadata As FontPropertyMetadata = TryCast([property].GetMetadata(GetType(FontChooser)), FontPropertyMetadata)
            If metadata IsNot Nothing Then
                obj.SetValue(metadata.TargetProperty, Me.GetValue([property]))
            End If
        Next
    End Sub

    Private Sub ApplyPropertiesToObjectExcept(ByVal obj As DependencyObject, ByVal except As DependencyProperty)
        For Each [property] As DependencyProperty In FontProperties
            If [property] IsNot except Then
                Dim metadata As FontPropertyMetadata = TryCast([property].GetMetadata(GetType(FontChooser)), FontPropertyMetadata)
                If metadata IsNot Nothing Then
                    obj.SetValue(metadata.TargetProperty, Me.GetValue([property]))
                End If
            End If
        Next [property]
    End Sub

    ''' <summary>
    ''' Sample text used in the preview box and family and typeface samples tab.
    ''' </summary>
    Public Property PreviewSampleText() As String
        Get
            Return _previewSampleText
        End Get

        Set(ByVal value As String)
            Dim newValue As String = If(String.IsNullOrEmpty(value), _defaultSampleText, value)
            If newValue <> _previewSampleText Then
                _previewSampleText = newValue

                ' Update the preview text box.
                previewTextBox.Text = newValue

                ' The preview sample text is also used in the family and typeface samples tab.
                InvalidateTab(samplesTab)
            End If
        End Set
    End Property

#End Region

#Region "Dependency properties for typographic features"

    Public Shared ReadOnly StandardLigaturesProperty As DependencyProperty = RegisterTypographicProperty(Typography.StandardLigaturesProperty)
    Public Property StandardLigatures() As Boolean
        Get
            Return CBool(GetValue(StandardLigaturesProperty))
        End Get
        Set(ByVal value As Boolean)
            SetValue(StandardLigaturesProperty, value)
        End Set
    End Property

    Public Shared ReadOnly ContextualLigaturesProperty As DependencyProperty = RegisterTypographicProperty(Typography.ContextualLigaturesProperty)
    Public Property ContextualLigatures() As Boolean
        Get
            Return CBool(GetValue(ContextualLigaturesProperty))
        End Get
        Set(ByVal value As Boolean)
            SetValue(ContextualLigaturesProperty, value)
        End Set
    End Property

    Public Shared ReadOnly DiscretionaryLigaturesProperty As DependencyProperty = RegisterTypographicProperty(Typography.DiscretionaryLigaturesProperty)
    Public Property DiscretionaryLigatures() As Boolean
        Get
            Return CBool(GetValue(DiscretionaryLigaturesProperty))
        End Get
        Set(ByVal value As Boolean)
            SetValue(DiscretionaryLigaturesProperty, value)
        End Set
    End Property

    Public Shared ReadOnly HistoricalLigaturesProperty As DependencyProperty = RegisterTypographicProperty(Typography.HistoricalLigaturesProperty)
    Public Property HistoricalLigatures() As Boolean
        Get
            Return CBool(GetValue(HistoricalLigaturesProperty))
        End Get
        Set(ByVal value As Boolean)
            SetValue(HistoricalLigaturesProperty, value)
        End Set
    End Property

    Public Shared ReadOnly ContextualAlternatesProperty As DependencyProperty = RegisterTypographicProperty(Typography.ContextualAlternatesProperty)
    Public Property ContextualAlternates() As Boolean
        Get
            Return CBool(GetValue(ContextualAlternatesProperty))
        End Get
        Set(ByVal value As Boolean)
            SetValue(ContextualAlternatesProperty, value)
        End Set
    End Property

    Public Shared ReadOnly HistoricalFormsProperty As DependencyProperty = RegisterTypographicProperty(Typography.HistoricalFormsProperty)
    Public Property HistoricalForms() As Boolean
        Get
            Return CBool(GetValue(HistoricalFormsProperty))
        End Get
        Set(ByVal value As Boolean)
            SetValue(HistoricalFormsProperty, value)
        End Set
    End Property

    Public Shared ReadOnly KerningProperty As DependencyProperty = RegisterTypographicProperty(Typography.KerningProperty)
    Public Property Kerning() As Boolean
        Get
            Return CBool(GetValue(KerningProperty))
        End Get
        Set(ByVal value As Boolean)
            SetValue(KerningProperty, value)
        End Set
    End Property

    Public Shared ReadOnly CapitalSpacingProperty As DependencyProperty = RegisterTypographicProperty(Typography.CapitalSpacingProperty)
    Public Property CapitalSpacing() As Boolean
        Get
            Return CBool(GetValue(CapitalSpacingProperty))
        End Get
        Set(ByVal value As Boolean)
            SetValue(CapitalSpacingProperty, value)
        End Set
    End Property

    Public Shared ReadOnly CaseSensitiveFormsProperty As DependencyProperty = RegisterTypographicProperty(Typography.CaseSensitiveFormsProperty)
    Public Property CaseSensitiveForms() As Boolean
        Get
            Return CBool(GetValue(CaseSensitiveFormsProperty))
        End Get
        Set(ByVal value As Boolean)
            SetValue(CaseSensitiveFormsProperty, value)
        End Set
    End Property

    Public Shared ReadOnly StylisticSet1Property As DependencyProperty = RegisterTypographicProperty(Typography.StylisticSet1Property)
    Public Property StylisticSet1() As Boolean
        Get
            Return CBool(GetValue(StylisticSet1Property))
        End Get
        Set(ByVal value As Boolean)
            SetValue(StylisticSet1Property, value)
        End Set
    End Property

    Public Shared ReadOnly StylisticSet2Property As DependencyProperty = RegisterTypographicProperty(Typography.StylisticSet2Property)
    Public Property StylisticSet2() As Boolean
        Get
            Return CBool(GetValue(StylisticSet2Property))
        End Get
        Set(ByVal value As Boolean)
            SetValue(StylisticSet2Property, value)
        End Set
    End Property

    Public Shared ReadOnly StylisticSet3Property As DependencyProperty = RegisterTypographicProperty(Typography.StylisticSet3Property)
    Public Property StylisticSet3() As Boolean
        Get
            Return CBool(GetValue(StylisticSet3Property))
        End Get
        Set(ByVal value As Boolean)
            SetValue(StylisticSet3Property, value)
        End Set
    End Property

    Public Shared ReadOnly StylisticSet4Property As DependencyProperty = RegisterTypographicProperty(Typography.StylisticSet4Property)
    Public Property StylisticSet4() As Boolean
        Get
            Return CBool(GetValue(StylisticSet4Property))
        End Get
        Set(ByVal value As Boolean)
            SetValue(StylisticSet4Property, value)
        End Set
    End Property

    Public Shared ReadOnly StylisticSet5Property As DependencyProperty = RegisterTypographicProperty(Typography.StylisticSet5Property)
    Public Property StylisticSet5() As Boolean
        Get
            Return CBool(GetValue(StylisticSet5Property))
        End Get
        Set(ByVal value As Boolean)
            SetValue(StylisticSet5Property, value)
        End Set
    End Property

    Public Shared ReadOnly StylisticSet6Property As DependencyProperty = RegisterTypographicProperty(Typography.StylisticSet6Property)
    Public Property StylisticSet6() As Boolean
        Get
            Return CBool(GetValue(StylisticSet6Property))
        End Get
        Set(ByVal value As Boolean)
            SetValue(StylisticSet6Property, value)
        End Set
    End Property

    Public Shared ReadOnly StylisticSet7Property As DependencyProperty = RegisterTypographicProperty(Typography.StylisticSet7Property)
    Public Property StylisticSet7() As Boolean
        Get
            Return CBool(GetValue(StylisticSet7Property))
        End Get
        Set(ByVal value As Boolean)
            SetValue(StylisticSet7Property, value)
        End Set
    End Property

    Public Shared ReadOnly StylisticSet8Property As DependencyProperty = RegisterTypographicProperty(Typography.StylisticSet8Property)
    Public Property StylisticSet8() As Boolean
        Get
            Return CBool(GetValue(StylisticSet8Property))
        End Get
        Set(ByVal value As Boolean)
            SetValue(StylisticSet8Property, value)
        End Set
    End Property

    Public Shared ReadOnly StylisticSet9Property As DependencyProperty = RegisterTypographicProperty(Typography.StylisticSet9Property)
    Public Property StylisticSet9() As Boolean
        Get
            Return CBool(GetValue(StylisticSet9Property))
        End Get
        Set(ByVal value As Boolean)
            SetValue(StylisticSet9Property, value)
        End Set
    End Property

    Public Shared ReadOnly StylisticSet10Property As DependencyProperty = RegisterTypographicProperty(Typography.StylisticSet10Property)
    Public Property StylisticSet10() As Boolean
        Get
            Return CBool(GetValue(StylisticSet10Property))
        End Get
        Set(ByVal value As Boolean)
            SetValue(StylisticSet10Property, value)
        End Set
    End Property

    Public Shared ReadOnly StylisticSet11Property As DependencyProperty = RegisterTypographicProperty(Typography.StylisticSet11Property)
    Public Property StylisticSet11() As Boolean
        Get
            Return CBool(GetValue(StylisticSet11Property))
        End Get
        Set(ByVal value As Boolean)
            SetValue(StylisticSet11Property, value)
        End Set
    End Property

    Public Shared ReadOnly StylisticSet12Property As DependencyProperty = RegisterTypographicProperty(Typography.StylisticSet12Property)
    Public Property StylisticSet12() As Boolean
        Get
            Return CBool(GetValue(StylisticSet12Property))
        End Get
        Set(ByVal value As Boolean)
            SetValue(StylisticSet12Property, value)
        End Set
    End Property

    Public Shared ReadOnly StylisticSet13Property As DependencyProperty = RegisterTypographicProperty(Typography.StylisticSet13Property)
    Public Property StylisticSet13() As Boolean
        Get
            Return CBool(GetValue(StylisticSet13Property))
        End Get
        Set(ByVal value As Boolean)
            SetValue(StylisticSet13Property, value)
        End Set
    End Property

    Public Shared ReadOnly StylisticSet14Property As DependencyProperty = RegisterTypographicProperty(Typography.StylisticSet14Property)
    Public Property StylisticSet14() As Boolean
        Get
            Return CBool(GetValue(StylisticSet14Property))
        End Get
        Set(ByVal value As Boolean)
            SetValue(StylisticSet14Property, value)
        End Set
    End Property

    Public Shared ReadOnly StylisticSet15Property As DependencyProperty = RegisterTypographicProperty(Typography.StylisticSet15Property)
    Public Property StylisticSet15() As Boolean
        Get
            Return CBool(GetValue(StylisticSet15Property))
        End Get
        Set(ByVal value As Boolean)
            SetValue(StylisticSet15Property, value)
        End Set
    End Property

    Public Shared ReadOnly StylisticSet16Property As DependencyProperty = RegisterTypographicProperty(Typography.StylisticSet16Property)
    Public Property StylisticSet16() As Boolean
        Get
            Return CBool(GetValue(StylisticSet16Property))
        End Get
        Set(ByVal value As Boolean)
            SetValue(StylisticSet16Property, value)
        End Set
    End Property

    Public Shared ReadOnly StylisticSet17Property As DependencyProperty = RegisterTypographicProperty(Typography.StylisticSet17Property)
    Public Property StylisticSet17() As Boolean
        Get
            Return CBool(GetValue(StylisticSet17Property))
        End Get
        Set(ByVal value As Boolean)
            SetValue(StylisticSet17Property, value)
        End Set
    End Property

    Public Shared ReadOnly StylisticSet18Property As DependencyProperty = RegisterTypographicProperty(Typography.StylisticSet18Property)
    Public Property StylisticSet18() As Boolean
        Get
            Return CBool(GetValue(StylisticSet18Property))
        End Get
        Set(ByVal value As Boolean)
            SetValue(StylisticSet18Property, value)
        End Set
    End Property

    Public Shared ReadOnly StylisticSet19Property As DependencyProperty = RegisterTypographicProperty(Typography.StylisticSet19Property)
    Public Property StylisticSet19() As Boolean
        Get
            Return CBool(GetValue(StylisticSet19Property))
        End Get
        Set(ByVal value As Boolean)
            SetValue(StylisticSet19Property, value)
        End Set
    End Property

    Public Shared ReadOnly StylisticSet20Property As DependencyProperty = RegisterTypographicProperty(Typography.StylisticSet20Property)
    Public Property StylisticSet20() As Boolean
        Get
            Return CBool(GetValue(StylisticSet20Property))
        End Get
        Set(ByVal value As Boolean)
            SetValue(StylisticSet20Property, value)
        End Set
    End Property

    Public Shared ReadOnly SlashedZeroProperty As DependencyProperty = RegisterTypographicProperty(Typography.SlashedZeroProperty, "Digits")
    Public Property SlashedZero() As Boolean
        Get
            Return CBool(GetValue(SlashedZeroProperty))
        End Get
        Set(ByVal value As Boolean)
            SetValue(SlashedZeroProperty, value)
        End Set
    End Property

    Public Shared ReadOnly MathematicalGreekProperty As DependencyProperty = RegisterTypographicProperty(Typography.MathematicalGreekProperty)
    Public Property MathematicalGreek() As Boolean
        Get
            Return CBool(GetValue(MathematicalGreekProperty))
        End Get
        Set(ByVal value As Boolean)
            SetValue(MathematicalGreekProperty, value)
        End Set
    End Property

    Public Shared ReadOnly EastAsianExpertFormsProperty As DependencyProperty = RegisterTypographicProperty(Typography.EastAsianExpertFormsProperty)
    Public Property EastAsianExpertForms() As Boolean
        Get
            Return CBool(GetValue(EastAsianExpertFormsProperty))
        End Get
        Set(ByVal value As Boolean)
            SetValue(EastAsianExpertFormsProperty, value)
        End Set
    End Property

    Public Shared ReadOnly FractionProperty As DependencyProperty = RegisterTypographicProperty(Typography.FractionProperty, "OneHalf")
    Public Property Fraction() As FontFraction
        Get
            Return CType(GetValue(FractionProperty), FontFraction)
        End Get
        Set(ByVal value As FontFraction)
            SetValue(FractionProperty, value)
        End Set
    End Property

    Public Shared ReadOnly VariantsProperty As DependencyProperty = RegisterTypographicProperty(Typography.VariantsProperty)
    Public Property Variants() As FontVariants
        Get
            Return CType(GetValue(VariantsProperty), FontVariants)
        End Get
        Set(ByVal value As FontVariants)
            SetValue(VariantsProperty, value)
        End Set
    End Property

    Public Shared ReadOnly CapitalsProperty As DependencyProperty = RegisterTypographicProperty(Typography.CapitalsProperty)
    Public Property Capitals() As FontCapitals
        Get
            Return CType(GetValue(CapitalsProperty), FontCapitals)
        End Get
        Set(ByVal value As FontCapitals)
            SetValue(CapitalsProperty, value)
        End Set
    End Property

    Public Shared ReadOnly NumeralStyleProperty As DependencyProperty = RegisterTypographicProperty(Typography.NumeralStyleProperty, "Digits")
    Public Property NumeralStyle() As FontNumeralStyle
        Get
            Return CType(GetValue(NumeralStyleProperty), FontNumeralStyle)
        End Get
        Set(ByVal value As FontNumeralStyle)
            SetValue(NumeralStyleProperty, value)
        End Set
    End Property

    Public Shared ReadOnly NumeralAlignmentProperty As DependencyProperty = RegisterTypographicProperty(Typography.NumeralAlignmentProperty, "Digits")
    Public Property NumeralAlignment() As FontNumeralAlignment
        Get
            Return CType(GetValue(NumeralAlignmentProperty), FontNumeralAlignment)
        End Get
        Set(ByVal value As FontNumeralAlignment)
            SetValue(NumeralAlignmentProperty, value)
        End Set
    End Property

    Public Shared ReadOnly EastAsianWidthsProperty As DependencyProperty = RegisterTypographicProperty(Typography.EastAsianWidthsProperty)
    Public Property EastAsianWidths() As FontEastAsianWidths
        Get
            Return CType(GetValue(EastAsianWidthsProperty), FontEastAsianWidths)
        End Get
        Set(ByVal value As FontEastAsianWidths)
            SetValue(EastAsianWidthsProperty, value)
        End Set
    End Property

    Public Shared ReadOnly EastAsianLanguageProperty As DependencyProperty = RegisterTypographicProperty(Typography.EastAsianLanguageProperty)
    Public Property EastAsianLanguage() As FontEastAsianLanguage
        Get
            Return CType(GetValue(EastAsianLanguageProperty), FontEastAsianLanguage)
        End Get
        Set(ByVal value As FontEastAsianLanguage)
            SetValue(EastAsianLanguageProperty, value)
        End Set
    End Property

    Public Shared ReadOnly AnnotationAlternatesProperty As DependencyProperty = RegisterTypographicProperty(Typography.AnnotationAlternatesProperty)
    Public Property AnnotationAlternates() As Integer
        Get
            Return CInt(Fix(GetValue(AnnotationAlternatesProperty)))
        End Get
        Set(ByVal value As Integer)
            SetValue(AnnotationAlternatesProperty, value)
        End Set
    End Property

    Public Shared ReadOnly StandardSwashesProperty As DependencyProperty = RegisterTypographicProperty(Typography.StandardSwashesProperty)
    Public Property StandardSwashes() As Integer
        Get
            Return CInt(Fix(GetValue(StandardSwashesProperty)))
        End Get
        Set(ByVal value As Integer)
            SetValue(StandardSwashesProperty, value)
        End Set
    End Property

    Public Shared ReadOnly ContextualSwashesProperty As DependencyProperty = RegisterTypographicProperty(Typography.ContextualSwashesProperty)
    Public Property ContextualSwashes() As Integer
        Get
            Return CInt(Fix(GetValue(ContextualSwashesProperty)))
        End Get
        Set(ByVal value As Integer)
            SetValue(ContextualSwashesProperty, value)
        End Set
    End Property

    Public Shared ReadOnly StylisticAlternatesProperty As DependencyProperty = RegisterTypographicProperty(Typography.StylisticAlternatesProperty)
    Public Property StylisticAlternates() As Integer
        Get
            Return CInt(Fix(GetValue(StylisticAlternatesProperty)))
        End Get
        Set(ByVal value As Integer)
            SetValue(StylisticAlternatesProperty, value)
        End Set
    End Property

    Private Shared Sub TypographicPropertyChangedCallback(ByVal obj As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
        Dim chooser As FontChooser = TryCast(obj, FontChooser)
        If chooser IsNot Nothing Then
            chooser.InvalidatePreview()
        End If
    End Sub

#End Region

#Region "Other dependency properties"

    Public Shared ReadOnly SelectedFontFamilyProperty As DependencyProperty =
             RegisterFontProperty("SelectedFontFamily",
             TextBlock.FontFamilyProperty,
             New PropertyChangedCallback(AddressOf SelectedFontFamilyChangedCallback))

    Public Property SelectedFontFamily() As FontFamily
        Get
            Return TryCast(GetValue(SelectedFontFamilyProperty), FontFamily)
        End Get
        Set(ByVal value As FontFamily)
            SetValue(SelectedFontFamilyProperty, value)
        End Set
    End Property

    Private Shared Sub SelectedFontFamilyChangedCallback(ByVal obj As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
        CType(obj, FontChooser).OnSelectedFontFamilyChanged(TryCast(e.NewValue, FontFamily))
    End Sub

    Public Shared ReadOnly SelectedFontWeightProperty As DependencyProperty =
                    RegisterFontProperty("SelectedFontWeight",
                    TextBlock.FontWeightProperty,
                    New PropertyChangedCallback(AddressOf SelectedTypefaceChangedCallback))

    Public Property SelectedFontWeight() As FontWeight
        Get
            Return CType(GetValue(SelectedFontWeightProperty), FontWeight)
        End Get
        Set(ByVal value As FontWeight)
            SetValue(SelectedFontWeightProperty, value)
        End Set
    End Property

    Public Shared ReadOnly SelectedFontStyleProperty As DependencyProperty = RegisterFontProperty("SelectedFontStyle", TextBlock.FontStyleProperty, New PropertyChangedCallback(AddressOf SelectedTypefaceChangedCallback))
    Public Property SelectedFontStyle() As FontStyle
        Get
            Return CType(GetValue(SelectedFontStyleProperty), FontStyle)
        End Get
        Set(ByVal value As FontStyle)
            SetValue(SelectedFontStyleProperty, value)
        End Set
    End Property

    Public Shared ReadOnly SelectedFontStretchProperty As DependencyProperty = RegisterFontProperty("SelectedFontStretch", TextBlock.FontStretchProperty, New PropertyChangedCallback(AddressOf SelectedTypefaceChangedCallback))
    Public Property SelectedFontStretch() As FontStretch
        Get
            Return CType(GetValue(SelectedFontStretchProperty), FontStretch)
        End Get
        Set(ByVal value As FontStretch)
            SetValue(SelectedFontStretchProperty, value)
        End Set
    End Property

    Private Shared Sub SelectedTypefaceChangedCallback(ByVal obj As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
        CType(obj, FontChooser).InvalidateTypefaceListSelection()
    End Sub

    Public Property FontSizeInPoints() As Double
        Get
            Return SelectedFontSize * 72 / 96
        End Get

        Set(ByVal value As Double)
            SelectedFontSize = value * 96 / 72
        End Set
    End Property

    Public Shared ReadOnly SelectedFontSizeProperty As DependencyProperty = RegisterFontProperty("SelectedFontSize", TextBlock.FontSizeProperty, New PropertyChangedCallback(AddressOf SelectedFontSizeChangedCallback))

    Public Property SelectedFontSize() As Double
        Get
            Return CDbl(GetValue(SelectedFontSizeProperty))
        End Get

        Set(ByVal value As Double)
            SetValue(SelectedFontSizeProperty, value)
        End Set
    End Property

    Private Shared Sub SelectedFontSizeChangedCallback(ByVal obj As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
        CType(obj, FontChooser).OnSelectedFontSizeChanged(CDbl(e.NewValue))
    End Sub

    Public Shared ReadOnly SelectedTextDecorationsProperty As DependencyProperty = RegisterFontProperty("SelectedTextDecorations", TextBlock.TextDecorationsProperty, New PropertyChangedCallback(AddressOf SelectedTextDecorationsChangedCallback))
    Public Property SelectedTextDecorations() As TextDecorationCollection
        Get
            Return TryCast(GetValue(SelectedTextDecorationsProperty), TextDecorationCollection)
        End Get
        Set(ByVal value As TextDecorationCollection)
            SetValue(SelectedTextDecorationsProperty, value)
        End Set
    End Property
    Private Shared Sub SelectedTextDecorationsChangedCallback(ByVal obj As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
        Dim chooser As FontChooser = CType(obj, FontChooser)
        chooser.OnTextDecorationsChanged()
    End Sub

#End Region

#Region "Dependency property helper functions"

    ' Helper function for registering typographic dependency properties with property-specific sample text.
    Private Shared Function RegisterTypographicProperty(ByVal targetProperty As DependencyProperty, ByVal sampleTextTag As String) As DependencyProperty
        Dim t As Type = targetProperty.PropertyType

        Dim featurePage As TypographyFeaturePage = If(t Is GetType(Boolean), TypographyFeaturePage.BooleanFeaturePage, If(t Is GetType(Integer), TypographyFeaturePage.IntegerFeaturePage, New TypographyFeaturePage(t)))

        Return DependencyProperty.Register(targetProperty.Name, t, GetType(FontChooser), New TypographicPropertyMetadata(targetProperty.DefaultMetadata.DefaultValue, targetProperty, featurePage, sampleTextTag))
    End Function

    ' Helper function for registering typographic dependency properties with default sample text for the type.
    Private Shared Function RegisterTypographicProperty(ByVal targetProperty As DependencyProperty) As DependencyProperty
        Return RegisterTypographicProperty(targetProperty, Nothing)
    End Function

    ' Helper function for registering font chooser dependency properties other than typographic properties.
    Private Shared Function RegisterFontProperty(ByVal propertyName As String, ByVal targetProperty As DependencyProperty, ByVal changeCallback As PropertyChangedCallback) As DependencyProperty
        Return DependencyProperty.Register(propertyName, targetProperty.PropertyType, GetType(FontChooser), New FontPropertyMetadata(targetProperty.DefaultMetadata.DefaultValue, changeCallback, targetProperty))
    End Function

#End Region

#Region "Dependency property tables"

    ' Array of all font chooser dependency properties
    Private Shared ReadOnly FontProperties() As DependencyProperty =
        {StandardLigaturesProperty, ContextualLigaturesProperty,
         DiscretionaryLigaturesProperty, HistoricalLigaturesProperty,
         ContextualAlternatesProperty, HistoricalFormsProperty,
         KerningProperty, CapitalSpacingProperty,
         CaseSensitiveFormsProperty, StylisticSet1Property,
         StylisticSet2Property, StylisticSet3Property, StylisticSet4Property,
         StylisticSet5Property, StylisticSet6Property, StylisticSet7Property,
         StylisticSet8Property, StylisticSet9Property, StylisticSet10Property,
         StylisticSet11Property, StylisticSet12Property, StylisticSet13Property,
         StylisticSet14Property, StylisticSet15Property, StylisticSet16Property,
         StylisticSet17Property, StylisticSet18Property, StylisticSet19Property,
         StylisticSet20Property, SlashedZeroProperty, MathematicalGreekProperty,
         EastAsianExpertFormsProperty, FractionProperty, VariantsProperty,
         CapitalsProperty, NumeralStyleProperty, NumeralAlignmentProperty,
         EastAsianWidthsProperty, EastAsianLanguageProperty,
         AnnotationAlternatesProperty, StandardSwashesProperty,
         ContextualSwashesProperty, StylisticAlternatesProperty,
         SelectedFontFamilyProperty, SelectedFontWeightProperty,
         SelectedFontStyleProperty, SelectedFontStretchProperty,
         SelectedFontSizeProperty, SelectedTextDecorationsProperty}


#End Region

#Region "Property change handlers"

    ' Handle changes to the SelectedFontFamily property
    Private Sub OnSelectedFontFamilyChanged(ByVal family As FontFamily)
        ' If the family list is not valid do nothing for now. 
        ' We'll be called again after the list is initialized.
        If _familyListValid Then
            ' Select the family in the list; this will return null if the family is not in the list.
            fontFamilyList.Text = New FontFamilyListItem(family).ToString
            ScrollIntoView(fontFamilyList)

            ' The typeface list is no longer valid; update it in the background to improve responsiveness.
            InvalidateTypefaceList()
        End If
    End Sub

    ' Handle changes to the SelectedFontSize property
    Private Sub OnSelectedFontSizeChanged(ByVal FontSize As Double)
        ' Set the text box contents if it doesn't already match the current size.
        Dim textBoxValue As Double
        If (Not Double.TryParse(sizeList.Text, textBoxValue)) OrElse (Not Math.Abs(textBoxValue - FontSizeInPoints) < 0.01) Then
            sizeList.Text = FontSizeInPoints.ToString()
        End If

        ' Schedule background updates.
        InvalidateTab(typographyTab)
        InvalidatePreview()
    End Sub

    ' Handle changes to any of the text decoration properties.
    Private Sub OnTextDecorationsChanged()
        Dim underline As Boolean = False
        Dim baseline As Boolean = False
        Dim strikethrough As Boolean = False
        Dim overline As Boolean = False

        Dim textDecorations As TextDecorationCollection = SelectedTextDecorations
        If textDecorations IsNot Nothing Then
            For Each td As TextDecoration In textDecorations
                Select Case td.Location
                    Case TextDecorationLocation.Underline
                        underline = True
                    Case TextDecorationLocation.Baseline
                        baseline = True
                    Case TextDecorationLocation.Strikethrough
                        strikethrough = True
                    Case TextDecorationLocation.OverLine
                        overline = True
                End Select
            Next td
        End If

        underlineCheckBox.IsChecked = underline
        baselineCheckBox.IsChecked = baseline
        strikethroughCheckBox.IsChecked = strikethrough
        overlineCheckBox.IsChecked = overline

        ' Schedule background updates.
        InvalidateTab(typographyTab)
        InvalidatePreview()
    End Sub

#End Region

#Region "Background update logic"

    ' Schedule background initialization of the font famiy list.
    Private Sub InvalidateFontFamilyList()
        If _familyListValid Then
            InvalidateTypefaceList()

            fontFamilyList.Items.Clear()
            fontFamilyTextBox.Clear()
            _familyListValid = False

            ScheduleUpdate()
        End If
    End Sub

    ' Schedule background initialization of the typeface list.
    Private Sub InvalidateTypefaceList()
        If _typefaceListValid Then
            typefaceList.Items.Clear()
            _typefaceListValid = False

            ScheduleUpdate()
        End If
    End Sub

    ' Schedule background selection of the current typeface list item.
    Private Sub InvalidateTypefaceListSelection()
        If _typefaceListSelectionValid Then
            _typefaceListSelectionValid = False
            ScheduleUpdate()
        End If
    End Sub

    ' Mark a specific tab as invalid and schedule background initialization if necessary.
    Private Sub InvalidateTab(ByVal tab As TabItem)
        Dim tabState As TabState
        If _tabDictionary.TryGetValue(tab, tabState) Then
            If tabState.IsValid Then
                tabState.IsValid = False

                If tabControl.SelectedItem Is tab Then
                    ScheduleUpdate()
                End If
            End If
        End If
    End Sub

    ' Mark all the tabs as invalid and schedule background initialization of the current tab.
    Private Sub InvalidateTabs()
        For Each item As KeyValuePair(Of TabItem, TabState) In _tabDictionary
            item.Value.IsValid = False
        Next item

        ScheduleUpdate()
    End Sub

    ' Schedule background initialization of the preview control.
    Private Sub InvalidatePreview()
        If _previewValid Then
            _previewValid = False
            ScheduleUpdate()
        End If
    End Sub

    ' Schedule background initialization.
    Private Sub ScheduleUpdate()
        If Not _updatePending Then
            Dispatcher.BeginInvoke(DispatcherPriority.ContextIdle, New UpdateCallback(AddressOf OnUpdate))
            _updatePending = True
        End If
    End Sub

    ' Dispatcher callback that performs background initialization.
    Private Sub OnUpdate()
        _updatePending = False

        If Not _familyListValid Then
            ' Initialize the font family list.
            InitializeFontFamilyList()
            _familyListValid = True
            OnSelectedFontFamilyChanged(SelectedFontFamily)

            ' Defer any other initialization until later.
            ScheduleUpdate()
        ElseIf Not _typefaceListValid Then
            ' Initialize the typeface list.
            InitializeTypefaceList()
            _typefaceListValid = True

            ' Select the current typeface in the list.
            InitializeTypefaceListSelection()
            _typefaceListSelectionValid = True

            ' Defer any other initialization until later.
            ScheduleUpdate()
        ElseIf Not _typefaceListSelectionValid Then
            ' Select the current typeface in the list.
            InitializeTypefaceListSelection()
            _typefaceListSelectionValid = True

            ' Defer any other initialization until later.
            ScheduleUpdate()
        Else
            ' Perform any remaining initialization.
            Dim tab As TabState = CurrentTabState
            If tab IsNot Nothing AndAlso (Not tab.IsValid) Then
                ' Initialize the current tab.
                tab.InitializeTab()
                tab.IsValid = True
            End If
            If Not _previewValid Then
                ' Initialize the preview control.
                InitializePreview()
                _previewValid = True
            End If
        End If
    End Sub

#End Region

#Region "Content initialization"

    Private Sub InitializeFontFamilyList()

        Dim items As List(Of FontFamilyListItem)

        Dim familyCollection As ICollection(Of FontFamily) = FontFamilyCollection
        If familyCollection IsNot Nothing Then
            items = New List(Of FontFamilyListItem)
            For Each family As FontFamily In familyCollection
                If FontFamilyListItem.CanShowFont(family) Then items.Add(New FontFamilyListItem(family))
            Next family

            items.Sort()
            fontFamilyList.ItemsSource = items
        End If


    End Sub

    Private Sub InitializeTypefaceList()
        Dim family As FontFamily = SelectedFontFamily
        If family IsNot Nothing Then
            Dim faceCollection As ICollection(Of Typeface) = family.GetTypefaces()

            Dim items(faceCollection.Count - 1) As TypefaceListItem

            Dim i As Integer = 0

            For Each face As Typeface In faceCollection
                items(i) = New TypefaceListItem(face)
                i += 1
            Next face

            Array.Sort(Of TypefaceListItem)(items)

            For Each item As TypefaceListItem In items
                typefaceList.Items.Add(item)
                'If item.FontWeight = Me.SelectedFontWeight AndAlso item.FontStyle = Me.SelectedFontStyle AndAlso item.FontStretch = Me.SelectedFontStretch Then
                '    typefaceList.SelectedItem = item
                'End If
            Next

        End If
    End Sub

    Private Sub InitializeTypefaceListSelection()
        ' If the typeface list is not valid, do nothing for now.
        ' We'll be called again after the list is initialized.
        If _typefaceListValid Then
            Dim typeface As New Typeface(SelectedFontFamily, SelectedFontStyle, SelectedFontWeight, SelectedFontStretch)
            ExitTypefaceListSelectionChanged = True
            TypefaceListItem.Select(typefaceList, typeface)
            ExitTypefaceListSelectionChanged = False

            ' Schedule background updates.
            InvalidateTabs()
            InvalidatePreview()
        End If
    End Sub

    Private Sub InitializeFeatureList()
        Dim items(FontProperties.Length - 1) As TypographicFeatureListItem

        Dim count As Integer = 0

        For Each [property] As DependencyProperty In FontProperties
            If TypeOf [property].GetMetadata(GetType(FontChooser)) Is TypographicPropertyMetadata Then
                Dim displayName As String = LookupString([property].Name)
                items(count) = New TypographicFeatureListItem(displayName, [property])
                count += 1
            End If
        Next [property]

        Array.Sort(items, 0, count)

        For i As Integer = 0 To count - 1
            featureList.Items.Add(items(i))
        Next i
    End Sub

    Private Shared Function LookupString(ByVal tag As String) As String
        Return My.Resources.ResourceManager.GetString(tag, CultureInfo.CurrentUICulture)
    End Function

    Private ReadOnly Property CurrentTabState() As TabState
        Get
            Dim tab As TabState
            Return If(_tabDictionary.TryGetValue(TryCast(tabControl.SelectedItem, TabItem), tab), tab, Nothing)
        End Get
    End Property

    Private Sub InitializeSamplesTab()
        Dim selectedFamily As FontFamily = SelectedFontFamily

        Dim selectedFace As New Typeface(selectedFamily, SelectedFontStyle, SelectedFontWeight, SelectedFontStretch)

        fontFamilyNameRun.Text = FontFamilyListItem.GetDisplayName(selectedFamily)
        typefaceNameRun.Text = TypefaceListItem.GetDisplayName(selectedFace)

        ' Create FontFamily samples document.
        Dim doc As New FlowDocument()
        For Each face As Typeface In selectedFamily.GetTypefaces()
            Dim labelPara As New Paragraph(New Run(TypefaceListItem.GetDisplayName(face)))
            labelPara.Margin = New Thickness(0)
            doc.Blocks.Add(labelPara)

            Dim samplePara As New Paragraph(New Run(_previewSampleText))
            samplePara.FontFamily = selectedFamily
            samplePara.FontWeight = face.Weight
            samplePara.FontStyle = face.Style
            samplePara.FontStretch = face.Stretch
            samplePara.FontSize = 16.0
            samplePara.Margin = New Thickness(0, 0, 0, 8)
            doc.Blocks.Add(samplePara)
        Next face

        fontFamilySamples.Document = doc

        ' Create typeface samples document.
        doc = New FlowDocument()
        For Each size As Double In New Double() {9.0, 10.0, 11.0, 12.0, 13.0, 14.0, 15.0, 16.0, 17.0}
            Dim labelText As String = String.Format("{0} {1}", size, _pointsText)
            Dim labelPara As New Paragraph(New Run(labelText))
            labelPara.Margin = New Thickness(0)
            doc.Blocks.Add(labelPara)

            Dim samplePara As New Paragraph(New Run(_previewSampleText))
            samplePara.FontFamily = selectedFamily
            samplePara.FontWeight = selectedFace.Weight
            samplePara.FontStyle = selectedFace.Style
            samplePara.FontStretch = selectedFace.Stretch
            samplePara.FontSize = size
            samplePara.Margin = New Thickness(0, 0, 0, 8)
            doc.Blocks.Add(samplePara)
        Next size

        typefaceSamples.Document = doc
    End Sub

    Private Sub InitializeTypographyTab()
        If featureList.Items.IsEmpty Then
            InitializeFeatureList()
            featureList.SelectedIndex = 0

            AddHandler featureList.SelectionChanged, AddressOf featureList_SelectionChanged
        End If

        Dim chooserProperty As DependencyProperty = Nothing
        Dim featurePage As TypographyFeaturePage = Nothing

        Dim listItem As TypographicFeatureListItem = TryCast(featureList.SelectedItem, TypographicFeatureListItem)
        If listItem IsNot Nothing Then
            Dim metadata As TypographicPropertyMetadata = TryCast(listItem.ChooserProperty.GetMetadata(GetType(FontChooser)), TypographicPropertyMetadata)
            If metadata IsNot Nothing Then
                chooserProperty = listItem.ChooserProperty
                featurePage = metadata.FeaturePage
            End If
        End If

        InitializeFeaturePage(TypographyFeatureGrid, chooserProperty, featurePage)
    End Sub

    Private Sub InitializeFeaturePage(ByVal grid As Grid, ByVal chooserProperty As DependencyProperty, ByVal page As TypographyFeaturePage)
        If page Is Nothing Then
            grid.Children.Clear()
            grid.RowDefinitions.Clear()
        Else
            ' Get the property value and metadata.
            Dim value As Object = Me.GetValue(chooserProperty)
            Dim metadata As TypographicPropertyMetadata = CType(chooserProperty.GetMetadata(GetType(FontChooser)), TypographicPropertyMetadata)

            ' Look up the sample text.
            Dim sampleText As String = "Text sample  " & "نموذج"

            If page Is _currentFeaturePage Then
                ' Update the state of the controls.
                For i As Integer = 0 To page.Items.Length - 1
                    ' Check the radio button if it matches the current property value.
                    If page.Items(i).Value.Equals(value) Then
                        Dim radioButton As RadioButton = CType(grid.Children(i * 2), RadioButton)
                        radioButton.IsChecked = True
                        SelectFeatureRow(radioButton)
                        ScvFeature.ScrollToVerticalOffset(SelectionBoreder.Margin.Top)
                    End If

                    ' Apply properties to the sample text block.
                    Dim sample As TextBlock = CType(grid.Children(i * 2 + 1), TextBlock)
                    sample.Text = sampleText
                    ApplyPropertiesToObjectExcept(sample, chooserProperty)
                    sample.SetValue(metadata.TargetProperty, page.Items(i).Value)
                    sample.SetValue(TextBlock.FontSizeProperty, 22.0)
                Next i
            Else
                grid.Children.Clear()
                grid.RowDefinitions.Clear()

                ' Add row definitions.
                For i As Integer = 0 To page.Items.Length - 1
                    Dim row As New RowDefinition()
                    row.Height = GridLength.Auto
                    grid.RowDefinitions.Add(row)
                Next i

                ' Add the controls.
                For i As Integer = 0 To page.Items.Length - 1
                    Dim ItemTag As String = page.Items(i).Tag
                    Dim radioContent As New TextBlock(New Run(LookupString(ItemTag)))
                    radioContent.TextWrapping = TextWrapping.Wrap

                    ' Add the radio button.
                    Dim radioButton As New RadioButton()
                    radioButton.Name = ItemTag
                    radioButton.Content = radioContent
                    radioButton.Padding = New Thickness(10, 2.5, 0.0, 2.5)
                    radioButton.VerticalAlignment = VerticalAlignment.Center
                    System.Windows.Controls.Grid.SetRow(radioButton, i)
                    grid.Children.Add(radioButton)

                    ' Check the radio button if it matches the current property value.
                    If page.Items(i).Value.Equals(value) Then
                        radioButton.IsChecked = True

                        Try ' Refresh Grid
                            grid.Dispatcher.Invoke(DispatcherPriority.Render, Sub() Exit Sub)
                        Catch
                        End Try

                        SelectFeatureRow(radioButton)
                        ScvFeature.ScrollToVerticalOffset(SelectionBoreder.Margin.Top)
                    End If

                    ' Hook up the event.
                    AddHandler radioButton.Checked, AddressOf featureRadioButton_Checked

                    ' Add the block with sample text.
                    Dim sample As New TextBlock(New Run(sampleText))
                    sample.Padding = New Thickness(30, 2.5, 5, 2.5)
                    sample.TextWrapping = TextWrapping.WrapWithOverflow
                    sample.Tag = radioButton
                    ApplyPropertiesToObjectExcept(sample, chooserProperty)
                    sample.SetValue(metadata.TargetProperty, page.Items(i).Value)
                    sample.SetValue(TextBlock.FontSizeProperty, 22.0)
                    System.Windows.Controls.Grid.SetRow(sample, i)
                    System.Windows.Controls.Grid.SetColumn(sample, 1)
                    grid.Children.Add(sample)
                    AddHandler sample.PreviewMouseLeftButtonDown, AddressOf FeatureSampleText_Click
                Next i

                ' Add borders between rows.
                For i As Integer = 0 To page.Items.Length - 1
                    Dim border As New Border()
                    border.IsHitTestVisible = False
                    border.BorderThickness = New Thickness(0.0, 0.0, 0.0, 1.0)
                    border.BorderBrush = SystemColors.ControlLightBrush
                    System.Windows.Controls.Grid.SetRow(border, i)
                    System.Windows.Controls.Grid.SetColumnSpan(border, 2)
                    grid.Children.Add(border)
                Next i
            End If
        End If

        _currentFeature = chooserProperty
        _currentFeaturePage = page
    End Sub

    Dim SelectedRadioButton As RadioButton

    Sub SelectFeatureRow(Optional Rdo As RadioButton = Nothing)
        Dim ScrollIntoView As Boolean
        If Rdo Is Nothing Then
            Rdo = SelectedRadioButton
            ScrollIntoView = False
        Else
            SelectedRadioButton = Rdo
            ScrollIntoView = True
        End If

        Try
            SelectionBoreder.Height = Rdo.ActualHeight + 5
            SelectionBoreder.Width = TypographyBorder.ActualWidth - If(ScvFeature.ComputedVerticalScrollBarVisibility = Windows.Visibility.Visible, SystemParameters.VerticalScrollBarWidth, 0) - 10
            Dim P As Point
            P = Rdo.TransformToAncestor(TypographyBorder).Transform(P)
            Dim SelTop = P.Y - 2.5
            SelectionBoreder.Margin = New Thickness(5, SelTop, 0, 0)
            If ScrollIntoView Then
                If SelTop < 0 Then
                    ScvFeature.ScrollToVerticalOffset(ScvFeature.VerticalOffset + SelTop - 3)
                ElseIf SelTop + SelectionBoreder.Height > ScvFeature.ViewportHeight Then
                    ScvFeature.ScrollToVerticalOffset(ScvFeature.VerticalOffset + SelTop + SelectionBoreder.Height - ScvFeature.ViewportHeight + 3)
                End If
            End If

        Catch
        End Try
    End Sub

    Private Sub featureRadioButton_Checked(ByVal sender As Object, ByVal e As RoutedEventArgs)
        SelectFeatureRow(sender)

        If _currentFeature IsNot Nothing AndAlso _currentFeaturePage IsNot Nothing Then
            Dim ItemTag As String = CType(sender, RadioButton).Name

            For Each item As TypographyFeaturePage.Item In _currentFeaturePage.Items
                If item.Tag = ItemTag Then
                    Me.SetValue(_currentFeature, item.Value)
                End If
            Next item
        End If
    End Sub

    Private Sub AddTableRow(ByVal rowGroup As TableRowGroup, ByVal leftText As String, ByVal rightText As String)
        Dim row As New TableRow()

        row.Cells.Add(New TableCell(New Paragraph(New Run(leftText))))
        row.Cells.Add(New TableCell(New Paragraph(New Run(rightText))))

        rowGroup.Rows.Add(row)
    End Sub

    Private Sub AddTableRow(ByVal rowGroup As TableRowGroup, ByVal leftText As String, ByVal rightStrings As IDictionary(Of CultureInfo, String))
        Dim rightText As String = NameDictionaryHelper.GetDisplayName(rightStrings)
        AddTableRow(rowGroup, leftText, rightText)
    End Sub

    Private Sub InitializeDescriptiveTextTab()
        Dim selectedTypeface As New Typeface(SelectedFontFamily, SelectedFontStyle, SelectedFontWeight, SelectedFontStretch)

        Dim glyphTypeface As GlyphTypeface
        If selectedTypeface.TryGetGlyphTypeface(glyphTypeface) Then
            ' Create a table with two columns.
            Dim table As New Table()
            table.CellSpacing = 5
            Dim leftColumn As New TableColumn()
            leftColumn.Width = New GridLength(2.0, GridUnitType.Star)
            table.Columns.Add(leftColumn)
            Dim rightColumn As New TableColumn()
            rightColumn.Width = New GridLength(3.0, GridUnitType.Star)
            table.Columns.Add(rightColumn)

            Dim rowGroup As New TableRowGroup()
            AddTableRow(rowGroup, "Family:", glyphTypeface.FamilyNames)
            AddTableRow(rowGroup, "Face:", glyphTypeface.FaceNames)
            AddTableRow(rowGroup, "Description:", glyphTypeface.Descriptions)
            AddTableRow(rowGroup, "Version:", glyphTypeface.VersionStrings)
            AddTableRow(rowGroup, "Copyright:", glyphTypeface.Copyrights)
            AddTableRow(rowGroup, "Trademark:", glyphTypeface.Trademarks)
            AddTableRow(rowGroup, "Manufacturer:", glyphTypeface.ManufacturerNames)
            AddTableRow(rowGroup, "Designer:", glyphTypeface.DesignerNames)
            AddTableRow(rowGroup, "Designer URL:", glyphTypeface.DesignerUrls)
            AddTableRow(rowGroup, "Vendor URL:", glyphTypeface.VendorUrls)
            AddTableRow(rowGroup, "Win32 Family:", glyphTypeface.Win32FamilyNames)
            AddTableRow(rowGroup, "Win32 Face:", glyphTypeface.Win32FaceNames)

            Try
                AddTableRow(rowGroup, "Font File URI:", glyphTypeface.FontUri.ToString())
            Catch e1 As System.Security.SecurityException
                ' Font file URI is privileged information; just skip it if we don't have access.
            End Try

            table.RowGroups.Add(rowGroup)

            fontDescriptionBox.Document = New FlowDocument(table)

            fontLicenseBox.Text = NameDictionaryHelper.GetDisplayName(glyphTypeface.LicenseDescriptions)
        Else
            fontDescriptionBox.Document = New FlowDocument()
            fontLicenseBox.Text = String.Empty
        End If
    End Sub

    Private Sub InitializePreview()
        ApplyPropertiesToObject(previewTextBox)
    End Sub

#End Region

#Region "List box helpers"

    Dim ExitScrollIntoView As Boolean = False
    Private Sub ScrollIntoView(Cmb As ComboBox)
        If ExitScrollIntoView Then Return

        Dim I = Cmb.SelectedIndex
        If I > -1 Then Return

        Dim Sc = CType(Cmb.Template.FindName("DropDownScrollViewer", Cmb), ScrollViewer)
        If Sc Is Nothing Then Return

        ExitScrollIntoView = True
        Dim txt = Cmb.Text
        Dim Index = If(I = -1, FindNearstMatch(Cmb, False), I)
        Cmb.SelectedIndex = Index
        Dim Offset = Sc.VerticalOffset

        If Cmb.SelectedIndex <> I Then Cmb.SelectedIndex = I

        If Cmb.Text <> txt Then
            Cmb.Text = txt
            Dim Tb = CType(Cmb.Template.FindName("PART_EditableTextBox", Cmb), TextBox)
            Tb.SelectionStart = Tb.Text.Length
        End If

        Sc.ScrollToVerticalOffset(Offset)
        ExitScrollIntoView = False
    End Sub

    ' Logic to handle UP and DOWN arrow keys in the text box associated with a list.
    ' Behavior is similar to a Win32 combo box.
    Private Sub OnComboBoxPreviewKeyDown(ByVal listBox As Selector, ByVal e As KeyEventArgs)
        Select Case e.Key
            Case Key.OemPeriod
                Dim Cmb = TryCast(listBox, ComboBox)
                If Cmb IsNot Nothing Then
                    If Not Cmb.Text.Contains(".") Then
                        Dim Txt As TextBox = Cmb.Template.FindName("PART_EditableTextBox", Cmb)
                        If Txt IsNot Nothing Then
                            Txt.SelectedText = "."
                            Txt.SelectionLength = 0
                            Txt.SelectionStart += 1
                        End If
                    End If
                    e.Handled = True
                End If

            Case Key.Up
                ' Move up from the current position.
                MoveListPosition(listBox, -1)
                e.Handled = True

            Case Key.Down
                ' Move down from the current position, unless the item at the current position is
                ' not already selected in which case select it.
                If listBox.Items.CurrentPosition = listBox.SelectedIndex Then
                    MoveListPosition(listBox, +1)
                Else
                    MoveListPosition(listBox, 0)
                End If
                e.Handled = True
        End Select
    End Sub

    Private Sub MoveListPosition(ByVal listBox As ListBox, ByVal distance As Integer)

        Dim i As Integer = listBox.Items.CurrentPosition + distance
        If i >= 0 AndAlso i < listBox.Items.Count Then
            listBox.Items.MoveCurrentToPosition(i)
            listBox.SelectedIndex = i
            listBox.ScrollIntoView(listBox.Items(i))
        End If
    End Sub

#End Region

    Dim fontFamilyTextBox As TextBox

    Private Sub fontFamilyList_Loaded(sender As Object, e As RoutedEventArgs) Handles fontFamilyList.Loaded
        fontFamilyTextBox = fontFamilyList.Template.FindName("PART_EditableTextBox", fontFamilyList)
        'Dim T = Now
        ' Initialize the font family list and the current family.
        If Not _familyListValid Then
            InitializeFontFamilyList()
            _familyListValid = True
            OnSelectedFontFamilyChanged(SelectedFontFamily)
        End If
        fontFamilyList.IsDropDownOpen = True
        fontFamilyList.IsDropDownOpen = False
        'MsgBox("WR: " & Now.Subtract(T).TotalMilliseconds)
    End Sub

    Private Sub Cmb_GotFocus(sender As Object, e As RoutedEventArgs)
        CType(sender, ComboBox).IsDropDownOpen = True
    End Sub

    Private Sub sizeList_TextChanged(sender As Object, e As TextChangedEventArgs)
        If Not sizeList.IsDropDownOpen Then sizeList.IsDropDownOpen = True
        Dim FontSize As Double
        If Double.TryParse(sizeList.Text, FontSize) Then
            If Not Math.Abs(FontSize - FontSizeInPoints) < 0.01 Then
                FontSizeInPoints = FontSize
            End If
        End If
    End Sub

    Private Sub ComboBoxText_MouseWheel(sender As Object, e As MouseWheelEventArgs)
        If Not e.OriginalSource.GetType().Name.StartsWith("TextBox") Then Return

        Dim Cmb As ComboBox = sender
        Dim I = Cmb.SelectedIndex
        If I = -1 Then I = FindNearstMatch(Cmb, e.Delta < 0)

        If (e.Delta > 0 AndAlso I > 0) OrElse (e.Delta < 0 AndAlso I < Cmb.Items.Count - 1) Then
            Cmb.SelectedIndex = I - e.Delta / 120
        End If

    End Sub

    Private Function FindNearstMatch(Cmb As ComboBox, Increasing As Boolean) As Integer
        Dim txt = Cmb.Text
        If IsNumeric(txt) Then txt = Math.Ceiling(Val(txt))
        Dim Count = Cmb.Items.Count - 1
        Dim L = txt.Length
        For c = 0 To L - 1
            Dim X = txt.Substring(0, L - c)
            For i = 0 To Count
                If Cmb.Items(i).ToString.StartsWith(X) Then Return i - If(Increasing, 1, 0)
            Next
        Next
        Return -1
    End Function

    Dim FirstTime As Boolean = True
    Private Sub sizeList_DropDownOpened(sender As Object, e As EventArgs) Handles sizeList.DropDownOpened
        If FirstTime Then
            FirstTime = False
            Dim Pup = CType(sizeList.Template.FindName("PART_Popup", sizeList), Popup)
            Pup.IsOpen = True
            ScrollIntoView(sizeList)
        End If
    End Sub

    Private Sub sizeList_Loaded(sender As Object, e As RoutedEventArgs) Handles sizeList.Loaded
        sizeList.IsDropDownOpen = False
    End Sub

    Private Sub FontSizeBoxText_PreviewTextInput(sender As Object, e As TextCompositionEventArgs)
        If Not IsNumeric(e.Text) Then
            e.Handled = True
            Beep()
        End If
    End Sub

    Private Sub ScvFeature_ScrollChanged(sender As Object, e As ScrollChangedEventArgs)
        SelectFeatureRow()
    End Sub

    Private Sub FeatureSampleText_Click(sender As Object, e As MouseButtonEventArgs)
        Dim Txt = CType(sender, TextBlock)
        Dim Rdo = CType(Txt.Tag, RadioButton)
        Rdo.IsChecked = True
    End Sub

    Private Sub typefaceList_Loaded(sender As Object, e As RoutedEventArgs)
        Dim txt = CType(typefaceList.Template.FindName("PART_EditableTextBox", typefaceList), TextBox)
        txt.IsReadOnly = True
    End Sub

    Private Sub Cmb_PreviewMouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs)
        If Not e.OriginalSource.GetType().Name.StartsWith("TextBox") Then Return

        If e.Source Is typefaceList Then
            typefaceList.IsDropDownOpen = Not typefaceList.IsDropDownOpen
            e.Handled = True
            Dim txt = CType(typefaceList.Template.FindName("PART_EditableTextBox", typefaceList), TextBox)
            txt.Focus()
        Else
            Dim Cmb = TryCast(e.Source, ComboBox)
            If Cmb Is Nothing Then Return
            If Not Cmb.IsDropDownOpen Then Cmb.IsDropDownOpen = True
        End If
    End Sub
End Class
