Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Globalization
Imports System.IO
Imports System.Runtime.Serialization
Imports System.Security.Permissions
Imports System.Windows
Imports System.Windows.Markup
Imports System.Windows.Media
Imports System.Windows.Media.TextFormatting
Imports System.Xml

Namespace Microsoft.Nautilus.Text.Editor

    <Serializable>
    Public NotInheritable Class TextFormattingRunProperties
        Inherits TextRunProperties
        Implements ISerializable, IObjectReference

        <NonSerialized> Private _typeface As Typeface
        <NonSerialized> Private _size As Double?
        <NonSerialized> Private _hintingSize As Double?
        <NonSerialized> Private _foregroundBrush As Brush
        <NonSerialized> Private _backgroundBrush As Brush
        <NonSerialized> Private _textDecorations As TextDecorationCollection
        <NonSerialized> Private _textEffects As TextEffectCollection
        <NonSerialized> Private _cultureInfo As CultureInfo
        <NonSerialized> Private Shared ExistingProperties As New List(Of TextFormattingRunProperties)
        <NonSerialized> Private Shared EmptyProperties As New TextFormattingRunProperties
        <NonSerialized> <ThreadStatic> Private Shared EmptyTextEffectCollection As TextEffectCollection
        <NonSerialized> <ThreadStatic> Private Shared EmptyTextDecorationCollection As TextDecorationCollection
        <NonSerialized> <ThreadStatic> Private Shared DefaultTypeface As Typeface

        Public Overrides ReadOnly Property BackgroundBrush As Brush
            Get
                Return If(_backgroundBrush, Brushes.Transparent)
            End Get
        End Property

        Public Overrides ReadOnly Property CultureInfo As CultureInfo
            Get
                Return If(_cultureInfo, CultureInfo.CurrentCulture)
            End Get
        End Property

        Public Overrides ReadOnly Property FontHintingEmSize As Double
            Get
                Return If(_hintingSize, 0.0)
            End Get
        End Property

        Public Overrides ReadOnly Property FontRenderingEmSize As Double
            Get
                Return If(_size, 0.0)
            End Get
        End Property

        Public Overrides ReadOnly Property ForegroundBrush As Brush
            Get
                Return If(_foregroundBrush, Brushes.Transparent)
            End Get
        End Property

        Public Overrides ReadOnly Property TextDecorations As TextDecorationCollection
            Get
                Return If(_textDecorations, EmptyTextDecorationCollection)
            End Get
        End Property

        Public Overrides ReadOnly Property TextEffects As TextEffectCollection
            Get
                Return If(_textEffects, EmptyTextEffectCollection)
            End Get
        End Property

        Public Overrides ReadOnly Property Typeface As Typeface
            Get
                Return If(_typeface, DefaultTypeface)
            End Get
        End Property

        Public ReadOnly Property BackgroundBrushEmpty As Boolean
            Get
                Return _backgroundBrush Is Nothing
            End Get
        End Property

        Public ReadOnly Property CultureInfoEmpty As Boolean
            Get
                Return _cultureInfo Is Nothing
            End Get
        End Property

        Public ReadOnly Property FontHintingEmSizeEmpty As Boolean
            Get
                Return Not _hintingSize.HasValue
            End Get
        End Property

        Public ReadOnly Property FontRenderingEmSizeEmpty As Boolean
            Get
                Return Not _size.HasValue
            End Get
        End Property

        Public ReadOnly Property ForegroundBrushEmpty As Boolean
            Get
                Return _foregroundBrush Is Nothing
            End Get
        End Property

        Public ReadOnly Property TextDecorationsEmpty As Boolean
            Get
                Return _textDecorations Is Nothing
            End Get
        End Property

        Public ReadOnly Property TextEffectsEmpty As Boolean
            Get
                Return _textEffects Is Nothing
            End Get
        End Property

        Public ReadOnly Property TypefaceEmpty As Boolean
            Get
                Return _typeface Is Nothing
            End Get
        End Property

        Friend Sub New()
        End Sub

        Friend Sub New(info As SerializationInfo, context As StreamingContext)
            _foregroundBrush = CType(GetObjectFromSerializationInfo("ForegroundBrush", info), Brush)
            _backgroundBrush = CType(GetObjectFromSerializationInfo("BackgroundBrush", info), Brush)
            _size = CType(GetObjectFromSerializationInfo("FontRenderingSize", info), Double?)
            _hintingSize = CType(GetObjectFromSerializationInfo("FontHintingSize", info), Double?)
            _textDecorations = CType(GetObjectFromSerializationInfo("TextDecorations", info), TextDecorationCollection)
            _textEffects = CType(GetObjectFromSerializationInfo("TextEffects", info), TextEffectCollection)
            _cultureInfo = CType(GetObjectFromSerializationInfo("CultureInfo", info), CultureInfo)

            Dim fontFamily1 As FontFamily = CType(GetObjectFromSerializationInfo("FontFamily", info), FontFamily)
            If fontFamily1 IsNot Nothing Then
                Dim style As FontStyle = CType(GetObjectFromSerializationInfo("Typeface.Style", info), FontStyle)
                Dim weight As FontWeight = CType(GetObjectFromSerializationInfo("Typeface.Weight", info), FontWeight)
                Dim stretch As FontStretch = CType(GetObjectFromSerializationInfo("Typeface.Stretch", info), FontStretch)
                _typeface = New Typeface(fontFamily1, style, weight, stretch)
            End If
        End Sub

        Friend Sub New(foreground As Brush, background As Brush, typeface1 As Typeface, size As Double?, hintingSize As Double?, textDecorations1 As TextDecorationCollection, textEffects1 As TextEffectCollection, cultureInfo1 As CultureInfo)
            _foregroundBrush = foreground
            _backgroundBrush = background
            _typeface = typeface1
            _size = size
            _hintingSize = hintingSize
            _textDecorations = textDecorations1
            _textEffects = textEffects1
            _cultureInfo = cultureInfo1
        End Sub

        Friend Sub New(toCopy As TextFormattingRunProperties)
            _foregroundBrush = toCopy._foregroundBrush
            _backgroundBrush = toCopy._backgroundBrush
            _typeface = toCopy._typeface
            _size = toCopy._size
            _hintingSize = toCopy._hintingSize
            _textDecorations = toCopy._textDecorations
            _textEffects = toCopy._textEffects
            _cultureInfo = toCopy._cultureInfo
        End Sub

        <EditorBrowsable(EditorBrowsableState.Always)>
        Public Shared Function CreateTextFormattingRunProperties() As TextFormattingRunProperties
            Return FindOrCreateProperties(EmptyProperties)
        End Function

        <EditorBrowsable(EditorBrowsableState.Always)>
        Public Shared Function CreateTextFormattingRunProperties(typeface1 As Typeface, size As Double, foreground As Color) As TextFormattingRunProperties
            Return FindOrCreateProperties(New TextFormattingRunProperties(New SolidColorBrush(foreground), Nothing, typeface1, size, Nothing, Nothing, Nothing, Nothing))
        End Function

        <EditorBrowsable(EditorBrowsableState.Advanced)>
        Public Shared Function CreateTextFormattingRunProperties(foreground As Brush, background As Brush, typeface As Typeface, size As Double?, hintingSize As Double?, textDecorations As TextDecorationCollection, textEffects As TextEffectCollection, cultureInfo As CultureInfo) As TextFormattingRunProperties
            Return FindOrCreateProperties(New TextFormattingRunProperties(foreground, background, typeface, size, hintingSize, textDecorations, textEffects, cultureInfo))
        End Function

        Public Function ClearBackgroundBrush() As TextFormattingRunProperties
            Dim textFormattingRunProperties1 As New TextFormattingRunProperties(Me)
            textFormattingRunProperties1._backgroundBrush = Nothing
            Return FindOrCreateProperties(textFormattingRunProperties1)
        End Function

        Public Function ClearCultureInfo() As TextFormattingRunProperties
            Dim textFormattingRunProperties1 As New TextFormattingRunProperties(Me)
            textFormattingRunProperties1._cultureInfo = Nothing
            Return FindOrCreateProperties(textFormattingRunProperties1)
        End Function

        Public Function ClearFontHintingEmSize() As TextFormattingRunProperties
            Dim textFormattingRunProperties1 As New TextFormattingRunProperties(Me)
            textFormattingRunProperties1._hintingSize = Nothing
            Return FindOrCreateProperties(textFormattingRunProperties1)
        End Function

        Public Function ClearFontRenderingEmSize() As TextFormattingRunProperties
            Dim textFormattingRunProperties1 As New TextFormattingRunProperties(Me)
            textFormattingRunProperties1._size = Nothing
            Return FindOrCreateProperties(textFormattingRunProperties1)
        End Function

        Public Function ClearForegroundBrush() As TextFormattingRunProperties
            Dim textFormattingRunProperties1 As New TextFormattingRunProperties(Me)
            textFormattingRunProperties1._foregroundBrush = Nothing
            Return FindOrCreateProperties(textFormattingRunProperties1)
        End Function

        Public Function ClearTextDecorations() As TextFormattingRunProperties
            Dim textFormattingRunProperties1 As New TextFormattingRunProperties(Me)
            textFormattingRunProperties1._textDecorations = Nothing
            Return FindOrCreateProperties(textFormattingRunProperties1)
        End Function

        Public Function ClearTextEffects() As TextFormattingRunProperties
            Dim textFormattingRunProperties1 As New TextFormattingRunProperties(Me)
            textFormattingRunProperties1._textEffects = Nothing
            Return FindOrCreateProperties(textFormattingRunProperties1)
        End Function

        Public Function ClearTypeface() As TextFormattingRunProperties
            Dim textFormattingRunProperties1 As New TextFormattingRunProperties(Me)
            textFormattingRunProperties1._typeface = Nothing
            Return FindOrCreateProperties(textFormattingRunProperties1)
        End Function

        Public Function SetBackgroundBrush(brush1 As Brush) As TextFormattingRunProperties
            If brush1 Is Nothing Then
                Throw New ArgumentNullException("brush")
            End If

            Dim textFormattingRunProperties1 As New TextFormattingRunProperties(Me)
            textFormattingRunProperties1._backgroundBrush = brush1
            textFormattingRunProperties1.FreezeBackgroundBrush()
            Return FindOrCreateProperties(textFormattingRunProperties1)
        End Function

        Public Function SetBackground(background As Color) As TextFormattingRunProperties
            Return SetBackgroundBrush(New SolidColorBrush(background))
        End Function

        Public Function SetCultureInfo(cultureInfo1 As CultureInfo) As TextFormattingRunProperties
            If cultureInfo1 Is Nothing Then
                Throw New ArgumentNullException("cultureInfo")
            End If

            Dim textFormattingRunProperties1 As New TextFormattingRunProperties(Me)
            textFormattingRunProperties1._cultureInfo = cultureInfo1
            Return FindOrCreateProperties(textFormattingRunProperties1)
        End Function

        Public Function SetFontHintingEmSize(hintingSize As Double) As TextFormattingRunProperties
            If hintingSize < 0.0 Then
                Throw New ArgumentOutOfRangeException("hintingSize")
            End If

            Dim textFormattingRunProperties1 As New TextFormattingRunProperties(Me)
            textFormattingRunProperties1._hintingSize = hintingSize
            Return FindOrCreateProperties(textFormattingRunProperties1)
        End Function

        Public Function SetFontRenderingEmSize(renderingSize As Double) As TextFormattingRunProperties
            If renderingSize <= 0.0 Then
                Throw New ArgumentOutOfRangeException("renderingSize")
            End If

            Dim textFormattingRunProperties1 As New TextFormattingRunProperties(Me)
            textFormattingRunProperties1._size = renderingSize
            Return FindOrCreateProperties(textFormattingRunProperties1)
        End Function

        Public Function SetForegroundBrush(brush1 As Brush) As TextFormattingRunProperties
            If brush1 Is Nothing Then
                Throw New ArgumentNullException("brush")
            End If

            Dim textFormattingRunProperties1 As New TextFormattingRunProperties(Me)
            textFormattingRunProperties1._foregroundBrush = brush1
            textFormattingRunProperties1.FreezeForegroundBrush()
            Return FindOrCreateProperties(textFormattingRunProperties1)
        End Function

        Public Function SetForeground(foreground As Color) As TextFormattingRunProperties
            Return SetForegroundBrush(New SolidColorBrush(foreground))
        End Function

        Public Function SetTextDecorations(textDecorations1 As TextDecorationCollection) As TextFormattingRunProperties
            If textDecorations1 Is Nothing Then
                Throw New ArgumentNullException("textDecorations")
            End If

            Dim textFormattingRunProperties1 As New TextFormattingRunProperties(Me)
            textFormattingRunProperties1._textDecorations = textDecorations1
            textFormattingRunProperties1.FreezeTextDecorations()
            Return FindOrCreateProperties(textFormattingRunProperties1)
        End Function

        Public Function SetTextEffects(textEffects1 As TextEffectCollection) As TextFormattingRunProperties
            If textEffects1 Is Nothing Then
                Throw New ArgumentNullException("textEffects")
            End If

            Dim textFormattingRunProperties1 As New TextFormattingRunProperties(Me)
            textFormattingRunProperties1._textEffects = textEffects1
            textFormattingRunProperties1.FreezeTextEffects()
            Return FindOrCreateProperties(textFormattingRunProperties1)
        End Function

        Public Function SetTypeface(typeface1 As Typeface) As TextFormattingRunProperties
            If typeface1 Is Nothing Then
                Throw New ArgumentNullException("typeface")
            End If
            Dim textFormattingRunProperties1 As New TextFormattingRunProperties(Me)
            textFormattingRunProperties1._typeface = typeface1
            Return FindOrCreateProperties(textFormattingRunProperties1)
        End Function

        Public Function SameSize(other As TextFormattingRunProperties) As Boolean
            Dim size2 = other._size
            If size2.HasValue Then
                Return _size.HasValue AndAlso _size.Value = size2.Value
            Else
                Return _size Is Nothing
            End If
        End Function

        Public Shared Function MergeProperties(favored As TextFormattingRunProperties, other As TextFormattingRunProperties) As TextFormattingRunProperties
            If other Is Nothing Then Return favored
            If favored Is Nothing Then Return other

            Dim textFormattingRunProperties1 As New TextFormattingRunProperties
            If favored._typeface IsNot Nothing Then
                textFormattingRunProperties1._typeface = favored._typeface
            Else
                textFormattingRunProperties1._typeface = other._typeface
            End If

            If favored._foregroundBrush IsNot Nothing Then
                textFormattingRunProperties1._foregroundBrush = favored._foregroundBrush
            Else
                textFormattingRunProperties1._foregroundBrush = other._foregroundBrush
            End If

            If favored._backgroundBrush IsNot Nothing Then
                textFormattingRunProperties1._backgroundBrush = favored._backgroundBrush
            Else
                textFormattingRunProperties1._backgroundBrush = other._backgroundBrush
            End If

            If favored._cultureInfo IsNot Nothing Then
                textFormattingRunProperties1._cultureInfo = favored._cultureInfo
            Else
                textFormattingRunProperties1._cultureInfo = other._cultureInfo
            End If

            If favored._size.HasValue Then
                textFormattingRunProperties1._size = favored._size
            Else
                textFormattingRunProperties1._size = other._size
            End If

            If favored._hintingSize.HasValue Then
                textFormattingRunProperties1._hintingSize = favored._hintingSize
            Else
                textFormattingRunProperties1._hintingSize = other._hintingSize
            End If

            If favored._textDecorations IsNot Nothing Then
                textFormattingRunProperties1._textDecorations = favored._textDecorations
            ElseIf other._textDecorations IsNot Nothing Then
                textFormattingRunProperties1._textDecorations = other._textDecorations
            End If

            If favored._textEffects IsNot Nothing Then
                textFormattingRunProperties1._textEffects = New TextEffectCollection(favored._textEffects)
            ElseIf other._textEffects IsNot Nothing Then
                textFormattingRunProperties1._textEffects = New TextEffectCollection(other._textEffects)
            End If

            Return FindOrCreateProperties(textFormattingRunProperties1)
        End Function

        Private Shared Function BrushesEqual(brush As Brush, other As Brush) As Boolean
            If brush Is Nothing Then
                Return other Is Nothing
            End If

            Dim solidColorBrush1 As SolidColorBrush = TryCast(brush, SolidColorBrush)
            Dim solidColorBrush2 As SolidColorBrush = TryCast(other, SolidColorBrush)
            If solidColorBrush1 IsNot Nothing AndAlso solidColorBrush2 IsNot Nothing Then
                Return solidColorBrush1.Color = solidColorBrush2.Color
            End If

            Return brush.Equals(other)
        End Function

        Private Shared Function TypefacesEqual(typeface1 As Typeface, other As Typeface) As Boolean
            Return If(typeface1?.Equals(other), (other Is Nothing))
        End Function

        Friend Shared Function FindOrCreateProperties(properties As TextFormattingRunProperties) As TextFormattingRunProperties
            Dim textFormattingRunProperties1 As TextFormattingRunProperties = ExistingProperties.Find(Function(other As TextFormattingRunProperties) properties.IsEqual(other))
            If textFormattingRunProperties1 Is Nothing Then
                properties.FreezeEverything()
                ExistingProperties.Add(properties)
                Return properties
            End If

            If EmptyTextDecorationCollection Is Nothing Then
                EmptyTextDecorationCollection = New TextDecorationCollection
            End If

            If EmptyTextEffectCollection Is Nothing Then
                EmptyTextEffectCollection = New TextEffectCollection
            End If

            If DefaultTypeface Is Nothing Then
                DefaultTypeface = New Typeface("Consolas")
            End If

            Return textFormattingRunProperties1
        End Function

        Private Function IsEqual(other As TextFormattingRunProperties) As Boolean
            If _size = other._size Then
                Dim hintingSize As Double? = _hintingSize
                Dim hintingSize2 As Double? = other._hintingSize
                If hintingSize.GetValueOrDefault() = hintingSize2.GetValueOrDefault() AndAlso hintingSize.HasValue = hintingSize2.HasValue AndAlso TypefacesEqual(_typeface, other._typeface) AndAlso _cultureInfo Is other._cultureInfo AndAlso _textDecorations Is other._textDecorations AndAlso _textEffects Is other._textEffects AndAlso BrushesEqual(_foregroundBrush, other._foregroundBrush) Then
                    Return BrushesEqual(_backgroundBrush, other._backgroundBrush)
                End If
            End If
            Return False
        End Function

        Private Sub FreezeBackgroundBrush()
            If _backgroundBrush IsNot Nothing AndAlso _backgroundBrush.CanFreeze Then
                _backgroundBrush.Freeze()
            End If
        End Sub

        Private Sub FreezeEverything()
            FreezeForegroundBrush()
            FreezeBackgroundBrush()
            FreezeTextEffects()
            FreezeTextDecorations()
        End Sub

        Private Sub FreezeForegroundBrush()
            If _foregroundBrush IsNot Nothing AndAlso _foregroundBrush.CanFreeze Then
                _foregroundBrush.Freeze()
            End If
        End Sub

        Private Sub FreezeTextDecorations()
            If _textDecorations IsNot Nothing AndAlso _textDecorations.CanFreeze Then
                _textDecorations.Freeze()
            End If
        End Sub

        Private Sub FreezeTextEffects()
            If _textEffects IsNot Nothing AndAlso _textEffects.CanFreeze Then
                _textEffects.Freeze()
            End If
        End Sub

        <SecurityPermission(SecurityAction.Demand, SerializationFormatter:=True)>
        <SecurityPermission(SecurityAction.LinkDemand, Flags:=SecurityPermissionFlag.SerializationFormatter)>
        Public Sub GetObjectData(info As SerializationInfo, context As StreamingContext) Implements ISerializable.GetObjectData
            If info Is Nothing Then
                Throw New ArgumentNullException("info")
            End If

            info.AddValue("BackgroundBrush", If(BackgroundBrushEmpty, "null", XamlWriter.Save(BackgroundBrush)))
            info.AddValue("ForegroundBrush", If(ForegroundBrushEmpty, "null", XamlWriter.Save(ForegroundBrush)))
            info.AddValue("CultureInfo", If(CultureInfoEmpty, "null", XamlWriter.Save(CultureInfo)))
            info.AddValue("FontHintingSize", If(FontHintingEmSizeEmpty, "null", XamlWriter.Save(FontHintingEmSize)))
            info.AddValue("FontRenderingSize", If(FontRenderingEmSizeEmpty, "null", XamlWriter.Save(FontRenderingEmSize)))
            info.AddValue("TextDecorations", If(TextDecorationsEmpty, "null", XamlWriter.Save(TextDecorations)))
            info.AddValue("TextEffects", If(TextEffectsEmpty, "null", XamlWriter.Save(TextEffects)))
            info.AddValue("FontFamily", If(TypefaceEmpty, "null", XamlWriter.Save(Typeface.FontFamily)))

            If Not TypefaceEmpty Then
                info.AddValue("Typeface.Style", XamlWriter.Save(Typeface.Style))
                info.AddValue("Typeface.Weight", XamlWriter.Save(Typeface.Weight))
                info.AddValue("Typeface.Stretch", XamlWriter.Save(Typeface.Stretch))
            End If
        End Sub

        Private Function GetObjectFromSerializationInfo(name As String, info As SerializationInfo) As Object
            Dim [string] As String = info.GetString(name)
            If [string] = "null" Then
                Return Nothing
            End If
            Using input As New StringReader([string])
                Return XamlReader.Load(XmlReader.Create(input))
            End Using
        End Function

        <SecurityPermission(SecurityAction.LinkDemand, Flags:=SecurityPermissionFlag.SerializationFormatter)>
        Public Function GetRealObject(context As StreamingContext) As Object Implements IObjectReference.GetRealObject
            Return FindOrCreateProperties(Me)
        End Function
    End Class
End Namespace
