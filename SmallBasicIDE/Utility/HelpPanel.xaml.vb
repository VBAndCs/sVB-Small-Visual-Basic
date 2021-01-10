Imports Microsoft.SmallBasic.LanguageService
Imports System
Imports System.ComponentModel
Imports System.Linq
Imports System.Text
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Documents
Imports System.Windows.Markup
Imports System.Windows.Media.Imaging

Namespace Microsoft.SmallBasic.Utility
    Public Partial Class HelpPanel
        Inherits Grid
        Implements INotifyPropertyChanged, IComponentConnector

        Private _completionItemWrapper As CompletionItemWrapper
        Private _paramsNameStyle As Style
        Private _paramsDescriptionStyle As Style
        Private _usageStyle As Style
        Private _typeStyle As Style
        Private _codeExampleStyle As Style
        Private _membersStyle As Style

        Public Property CompletionItemWrapper As CompletionItemWrapper
            Get
                Return _completionItemWrapper
            End Get

            Set(ByVal value As CompletionItemWrapper)
                _completionItemWrapper = value
                DataContext = _completionItemWrapper
                SetSymbolImage()
                Me.detailsParagraph.Inlines.Clear()

                If _completionItemWrapper.SymbolType = LanguageService.SymbolType.Method Then
                    FillArguments()
                ElseIf _completionItemWrapper.SymbolType = LanguageService.symbolType.Type Then
                    FillTypeMembers()
                End If

                FillExample()
            End Set
        End Property

        Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

        Public Sub New()
            Me.InitializeComponent()
            _paramsNameStyle = CType(FindResource("paramsNameStyle"), Style)
            _paramsDescriptionStyle = CType(FindResource("paramsDescriptionStyle"), Style)
            _usageStyle = CType(FindResource("usageStyle"), Style)
            _typeStyle = CType(FindResource("typeStyle"), Style)
            _codeExampleStyle = CType(FindResource("codeExampleStyle"), Style)
            _membersStyle = CType(FindResource("membersStyle"), Style)
        End Sub

        Private Sub SetSymbolImage()
            Select Case CompletionItemWrapper.SymbolType
                Case LanguageService.SymbolType.Event
                    Me.symbolType.Text = ResourceHelper.GetString("Event")
                Case LanguageService.SymbolType.Keyword
                    Me.symbolType.Text = ResourceHelper.GetString("Keyword")
                Case LanguageService.SymbolType.Method
                    Me.symbolType.Text = ResourceHelper.GetString("Operation")
                Case LanguageService.SymbolType.Property
                    Me.symbolType.Text = ResourceHelper.GetString("Property")
                Case LanguageService.SymbolType.Type
                    Me.symbolType.Text = ResourceHelper.GetString("Object")
                Case LanguageService.SymbolType.Variable
                    Me.symbolType.Text = ResourceHelper.GetString("Variable")
                Case LanguageService.SymbolType.Subroutine
                    Me.symbolType.Text = ResourceHelper.GetString("Subroutine")
                Case LanguageService.SymbolType.Label
                    Me.symbolType.Text = ResourceHelper.GetString("Label")
                Case Else
                    Me.symbolType.Text = ResourceHelper.GetString("Other")
            End Select

            Me.symbolImage.Source = GetIconForSymbolType(CompletionItemWrapper.SymbolType)
        End Sub

        Private Function GetIconForSymbolType(ByVal symbolType As SymbolType) As BitmapImage
            Select Case symbolType
                Case SymbolType.Event
                    Return New BitmapImage(New Uri("pack://application:,,/SB;component/Resources/IntellisenseEvent.png"))
                Case SymbolType.Keyword
                    Return New BitmapImage(New Uri("pack://application:,,/SB;component/Resources/IntellisenseKeyword.png"))
                Case SymbolType.Method
                    Return New BitmapImage(New Uri("pack://application:,,/SB;component/Resources/IntellisenseMethod.png"))
                Case SymbolType.Property
                    Return New BitmapImage(New Uri("pack://application:,,/SB;component/Resources/IntellisenseProperty.png"))
                Case SymbolType.Type
                    Return New BitmapImage(New Uri("pack://application:,,/SB;component/Resources/IntellisenseItem.png"))
                Case SymbolType.Variable
                    Return New BitmapImage(New Uri("pack://application:,,/SB;component/Resources/IntellisenseVariable.png"))
                Case SymbolType.Subroutine
                    Return New BitmapImage(New Uri("pack://application:,,/SB;component/Resources/IntellisenseSubroutine.png"))
                Case SymbolType.Label
                    Return New BitmapImage(New Uri("pack://application:,,/SB;component/Resources/IntellisenseLabel.png"))
                Case Else
                    Return New BitmapImage(New Uri("pack://application:,,/SB;component/Resources/IntellisenseItem.png"))
            End Select
        End Function

        Private Sub FillArguments()
            If CompletionItemWrapper.Documentation IsNot Nothing Then
                Dim name = CompletionItemWrapper.CompletionItem.MemberInfo.DeclaringType.Name
                Dim stringBuilder As StringBuilder = New StringBuilder($"{CompletionItemWrapper.Display}(")
                Dim flag = False

                For Each key In CompletionItemWrapper.Documentation.ParamsDoc.Keys
                    stringBuilder.Append(key & ", ")
                    flag = True
                Next

                If flag Then
                    stringBuilder.Remove(stringBuilder.Length - 2, 2)
                End If

                stringBuilder.Append(")")
                Dim span As Span = New Span()
                span.FlowDirection = FlowDirection.LeftToRight
                Dim span2 = span
                span2.Inlines.Add(New TextRun With {
                    .Text = name & ".",
                    .Style = _typeStyle
                })
                span2.Inlines.Add(New LineBreak())
                span2.Inlines.Add(New TextRun With {
                    .Text = stringBuilder.ToString(),
                    .Style = _usageStyle
                })
                span2.Inlines.Add(New LineBreak())
                span2.Inlines.Add(New LineBreak())

                For Each key2 In CompletionItemWrapper.Documentation.ParamsDoc.Keys
                    Dim textRun As TextRun = New TextRun()
                    textRun.Text = key2
                    textRun.Style = _paramsNameStyle
                    Dim item = textRun
                    span2.Inlines.Add(item)
                    span2.Inlines.Add(New LineBreak())
                    Dim textRun2 As TextRun = New TextRun()
                    textRun2.Text = CompletionItemWrapper.Documentation.ParamsDoc(key2)
                    textRun2.Style = _paramsDescriptionStyle
                    item = textRun2
                    span2.Inlines.Add(item)
                    span2.Inlines.Add(New LineBreak())
                Next

                span2.Inlines.Add(New Separator With {
                    .Width = 800.0,
                    .Margin = New Thickness(0.0, 4.0, 0.0, 4.0)
                })
                span2.Inlines.Add(New TextRun With {
                    .Text = ResourceHelper.GetString("Returns"),
                    .Style = _paramsNameStyle
                })
                span2.Inlines.Add(New LineBreak())

                If Equals(CompletionItemWrapper.Documentation.Returns, Nothing) OrElse CompletionItemWrapper.Documentation.Returns.Length = 0 Then
                    span2.Inlines.Add(New TextRun With {
                        .Text = ResourceHelper.GetString("Nothing"),
                        .Style = _paramsDescriptionStyle
                    })
                Else
                    span2.Inlines.Add(New TextRun With {
                        .Text = CompletionItemWrapper.Documentation.Returns,
                        .Style = _paramsDescriptionStyle
                    })
                End If

                span2.Inlines.Add(New LineBreak())
                Me.detailsParagraph.Inlines.Add(span2)
            End If
        End Sub

        Private Sub FillTypeMembers()
            Dim value As TypeInfo = Nothing

            If CompletionItemWrapper.Documentation IsNot Nothing Then
                Dim name = CompletionItemWrapper.CompletionItem.MemberInfo.Name

                If CompilerService.DummyCompiler.TypeInfoBag.Types.TryGetValue(name.ToLowerInvariant(), value) Then
                    For Each item In value.Properties.OrderBy(Function(ByVal a) a.Key)
                        Me.detailsParagraph.Inlines.Add(New Image With {
                            .Source = GetIconForSymbolType(LanguageService.SymbolType.Property),
                            .Width = 16.0,
                            .Height = 16.0,
                            .Margin = New Thickness(2.0),
                            .VerticalAlignment = VerticalAlignment.Center
                        })
                        Me.detailsParagraph.Inlines.Add(New TextRun With {
                            .Text = item.Value.Name & VisualBasic.Constants.vbCrLf,
                            .Style = _membersStyle
                        })
                    Next

                    For Each item2 In value.Methods.OrderBy(Function(ByVal a) a.Key)
                        Me.detailsParagraph.Inlines.Add(New Image With {
                            .Source = GetIconForSymbolType(LanguageService.SymbolType.Method),
                            .Width = 16.0,
                            .Height = 16.0,
                            .Margin = New Thickness(2.0),
                            .VerticalAlignment = VerticalAlignment.Center
                        })
                        Me.detailsParagraph.Inlines.Add(New TextRun With {
                            .Text = item2.Value.Name & VisualBasic.Constants.vbCrLf,
                            .Style = _membersStyle
                        })
                    Next

                    For Each item3 In value.Events.OrderBy(Function(ByVal a) a.Key)
                        Me.detailsParagraph.Inlines.Add(New Image With {
                            .Source = GetIconForSymbolType(LanguageService.SymbolType.Event),
                            .Width = 16.0,
                            .Height = 16.0,
                            .Margin = New Thickness(2.0),
                            .VerticalAlignment = VerticalAlignment.Center
                        })
                        Me.detailsParagraph.Inlines.Add(New TextRun With {
                            .Text = item3.Value.Name & VisualBasic.Constants.vbCrLf,
                            .Style = _membersStyle
                        })
                    Next
                End If
            End If
        End Sub

        Private Sub FillExample()
            If CompletionItemWrapper.Documentation IsNot Nothing AndAlso Not Equals(CompletionItemWrapper.Documentation.Example, Nothing) Then
                Me.detailsParagraph.Inlines.Add(New TextRun With {
                    .Text = ResourceHelper.GetString("Example"),
                    .Style = _paramsNameStyle
                })
                Me.detailsParagraph.Inlines.Add(New LineBreak())
                '__ = CompletionItemWrapper.Documentation.Example
                Me.detailsParagraph.Inlines.Add(New TextRun With {
                    .Text = CompletionItemWrapper.Documentation.Example,
                    .Style = _codeExampleStyle
                })
                Me.detailsParagraph.Inlines.Add(New LineBreak())
            End If
        End Sub

        Private Sub Notify(ByVal [property] As String)
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs([property]))
        End Sub
    End Class
End Namespace
