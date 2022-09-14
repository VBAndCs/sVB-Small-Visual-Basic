Imports Microsoft.SmallBasic.Completion
Imports Microsoft.SmallBasic.LanguageService
Imports Microsoft.VisualBasic
Imports System
Imports System.ComponentModel
Imports System.Linq
Imports System.Text
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Controls.Primitives
Imports System.Windows.Documents
Imports System.Windows.Markup
Imports System.Windows.Media.Imaging
Imports System.Windows.Navigation

Namespace Microsoft.SmallBasic.Utility
    Partial Public Class HelpPanel
        Inherits Grid
        Implements INotifyPropertyChanged, IComponentConnector

        Private _itemWrapper As CompletionItemWrapper
        Private _paramsNameStyle As Style
        Private _paramsDescriptionStyle As Style
        Private _usageStyle As Style
        Private _typeStyle As Style
        Private _selectedStyle As Style
        Private _codeExampleStyle As Style
        Private _membersStyle As Style
        Private _linkStyle As Style

        Dim gotoRun As New TextRun()
        Dim WithEvents gotoDefinintion As New Hyperlink(gotoRun)

        Public Property CompletionItemWrapper As CompletionItemWrapper
            Get
                Return _itemWrapper
            End Get

            Set(value As CompletionItemWrapper)
                _itemWrapper = value
                Dim item = value.CompletionItem


                Dim doc = MainWindow.ActiveDocument
                Dim textView = CType(doc.EditorControl.TextView, Nautilus.Text.Editor.AvalonTextView)
                Dim popHelp = MainWindow.popHelp

                If App.CntxMenuIsOpened OrElse CompletionItemWrapper.Documentation Is Nothing Then
                    popHelp.IsOpen = False
                    Return
                End If

                helpDocument.Blocks.Clear()
                methodType = _itemWrapper.Documentation?.Suffix

                Select Case value.SymbolType
                    Case SymbolType.Method, SymbolType.Subroutine
                        FillInfo(True)

                    Case SymbolType.DynamicProperty
                        methodType = " Dynamic Property"
                        FillInfo(False)

                    Case SymbolType.Control, SymbolType.Label, SymbolType.Literal,
                        SymbolType.GlobalVariable, SymbolType.LocalVariable
                        FillInfo(False)

                    Case SymbolType.Property, SymbolType.Event, SymbolType.Type
                        methodType = " " & value.SymbolType.ToString()
                        Dim type = GetTypeName()
                        Dim definition As New Span()
                        FillTitle(definition)
                        FillDescription(definition, If(type = "", "", type & "."))

                    Case Else
                        popHelp.IsOpen = False
                        Return
                End Select

                FillExample()

                popHelp.PlacementTarget = textView
                Dim caret = textView.Caret
                popHelp.HorizontalOffset = 10
                popHelp.MaxWidth = textView.ActualWidth - 20
                popHelp.MaxHeight = Math.Min(250, textView.ActualHeight - caret.Top)
                popHelp.VerticalOffset = caret.Top + caret.Height + If(IsCompletionListOpen(), 140, 0) + 5
                popHelp.IsOpen = True
                popHelp.Tag = textView
                AddHandler textView.ScrollChaged, AddressOf OnScrollChaged
            End Set
        End Property

        Private Function GetTypeName() As String
            Dim type = _itemWrapper.CompletionItem.MemberInfo?.DeclaringType
            If type Is Nothing Then Return ""

            Select Case type.Name
                Case NameOf(WinForms.ArrayEx)
                    Return NameOf(Library.Array)
                Case NameOf(WinForms.MathEx)
                    Return NameOf(Library.Math)
                Case NameOf(WinForms.TextEx)
                    Return NameOf(Library.Text)
                Case NameOf(WinForms.ColorEx)
                    Return NameOf(WinForms.Color)
                Case Else
                    Return type.Name
            End Select
        End Function

        Private Sub OnScrollChaged(senmder As Object, e As ScrollEventArgs)
            Dim popHelp = MainWindow.PopHelp
            popHelp.IsOpen = False
            Dim textView = CType(popHelp.Tag, Microsoft.Nautilus.Text.Editor.IAvalonTextView)
            If textView IsNot Nothing Then
                RemoveHandler textView.ScrollChaged, AddressOf OnScrollChaged
            End If
        End Sub

        Private Sub FillDescription(memberName As Span, belongsTo As String, Optional showSummery As Boolean = True)
            Dim documentation = CompletionItemWrapper.Documentation

            Dim p = New Paragraph()
            helpDocument.Blocks.Add(p)
            p.Inlines.Add(New TextRun With {
                    .Text = belongsTo,
                    .Style = _typeStyle
                }
            )

            p.Inlines.Add(memberName)

            If showSummery AndAlso documentation.Summary <> "" Then
                helpDocument.Blocks.Add(
                    New Paragraph(
                        New TextRun With {
                            .Text = documentation.Summary,
                            .Style = _paramsDescriptionStyle
                        }
                    ) With {.Margin = New Thickness(15, 0, 0, 5)}
                )
            End If
        End Sub

        Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

        Dim _mainWindow As MainWindow
        Private methodType As String

        Private ReadOnly Property MainWindow As MainWindow
            Get
                If _mainWindow Is Nothing Then
                    _mainWindow = Application.Current.MainWindow
                End If
                Return _mainWindow
            End Get
        End Property

        Function IsCompletionListOpen() As Boolean
            Dim completionSurface As CompletionAdornmentSurface = Nothing
            Dim properties = MainWindow.ActiveDocument.EditorControl.TextView.Properties
            Dim adornmentType = GetType(CompletionAdornmentSurface)

            Return properties.TryGetProperty(adornmentType, completionSurface) AndAlso
                    completionSurface.IsAdornmentVisible

        End Function

        Public Sub New()
            Me.InitializeComponent()
            _paramsNameStyle = CType(FindResource("paramsNameStyle"), Style)
            _paramsDescriptionStyle = CType(FindResource("paramsDescriptionStyle"), Style)
            _usageStyle = CType(FindResource("usageStyle"), Style)
            _typeStyle = CType(FindResource("typeStyle"), Style)
            _selectedStyle = CType(FindResource("selectedStyle"), Style)
            _linkStyle = CType(FindResource("linkStyle"), Style)
            _codeExampleStyle = CType(FindResource("codeExampleStyle"), Style)
            _membersStyle = CType(FindResource("membersStyle"), Style)
            gotoDefinintion.NavigateUri = New Uri("___test", UriKind.RelativeOrAbsolute)
        End Sub

        Private Function GetIconForSymbolType(symbolType As SymbolType) As BitmapImage
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
                Case SymbolType.GlobalVariable
                    Return New BitmapImage(New Uri("pack://application:,,/SB;component/Resources/IntellisenseVariable.png"))
                Case SymbolType.Subroutine
                    Return New BitmapImage(New Uri("pack://application:,,/SB;component/Resources/IntellisenseSubroutine.png"))
                Case SymbolType.Label
                    Return New BitmapImage(New Uri("pack://application:,,/SB;component/Resources/IntellisenseLabel.png"))
                Case Else
                    Return New BitmapImage(New Uri("pack://application:,,/SB;component/Resources/IntellisenseItem.png"))
            End Select
        End Function

        Private Sub FillInfo(hasParams As Boolean)
            Dim documentation = _itemWrapper.Documentation
            Dim item = _itemWrapper.CompletionItem
            Dim name = GetTypeName()
            Dim definition As New Span()
            Dim params = documentation.ParamsDoc.Keys
            Dim paramIndex = item.ParamIndex
            Dim addReturnOnly = paramIndex >= params.Count
            Dim addAllParams = paramIndex = -1 OrElse (addReturnOnly AndAlso documentation.Returns = "")

            FillTitle(definition)

            If hasParams Then
                definition.Inlines.Add(New TextRun() With {
                        .Text = "(",
                        .Style = _typeStyle
                 })

                Dim n = params.Count - 1
                For i = 0 To n
                    definition.Inlines.Add(New TextRun() With {
                        .Text = params(i),
                        .Style = If(i = paramIndex, _selectedStyle, _typeStyle)
                     })

                    If i < n Then
                        definition.Inlines.Add(New TextRun() With {
                             .Text = ", ",
                            .Style = _typeStyle
                        })
                    End If
                Next

                definition.Inlines.Add(New TextRun() With {
                        .Text = ")",
                        .Style = _typeStyle
                 })
            End If

            FillDescription(definition, If(name = "", documentation.Prefix, name & "."), addAllParams)

            ' Add param Info
            If params.Count > 0 Then

                Dim paramsDocs As New Paragraph With {
                    .Margin = New Thickness(15, 0, 0, 5)
                }

                Dim inlines = paramsDocs.Inlines
                Dim n = params.Count - 1

                For i = 0 To n
                    If Not addAllParams AndAlso i <> paramIndex Then Continue For

                    Dim paramDoc = documentation.ParamsDoc(params(i))
                    If paramDoc <> "" Then
                        inlines.Add(New TextRun() With {
                                .Text = params(i) & ": ",
                                .Style = _paramsNameStyle
                            })

                        inlines.Add(New TextRun() With {
                                .Text = paramDoc & If(addAllParams AndAlso i < n, vbCrLf, ""),
                                .Style = _paramsDescriptionStyle
                            })
                    End If
                Next

                If inlines.Count > 0 Then
                    If addAllParams Then
                        helpDocument.Blocks.Add(
                            New Paragraph(New TextRun With {
                                .Text = $"Parameter{If(addAllParams AndAlso n > 1, "s", "")} Info: ",
                                .Style = _paramsNameStyle
                            })
                        )
                    End If

                    helpDocument.Blocks.Add(paramsDocs)
                End If
            End If

            ' Add Return info
            If (addAllParams OrElse addReturnOnly) AndAlso documentation.Returns <> "" Then
                Dim p = New Paragraph() With {.Margin = New Thickness(0, 0, 0, 5)}
                helpDocument.Blocks.Add(p)
                Dim inlines = p.Inlines
                inlines.Add(New TextRun With {
                        .Text = ResourceHelper.GetString("Returns") & ": ",
                        .Style = _paramsNameStyle
                    })

                inlines.Add(New TextRun With {
                        .Text = documentation.Returns,
                        .Style = _paramsDescriptionStyle
                    })
            End If
        End Sub

        Private Sub FillTitle(definition As Span)
            Dim item = _itemWrapper.CompletionItem

            If Not item.DefinitionIdintifier.IsIllegal Then
                gotoRun.Text = _itemWrapper.Display
                gotoRun.Style = _linkStyle
                definition.Inlines.Add(gotoDefinintion)
                definition.Inlines.Add(New TextRun() With {
                    .Text = methodType,
                    .Style = _typeStyle
                })
            Else
                definition.Inlines.Add(New TextRun() With {
                    .Text = _itemWrapper.Display & methodType,
                    .Style = _typeStyle
                })
            End If
        End Sub

        Private Sub FillTypeMembers()
            Dim value As TypeInfo = Nothing
            Dim detailsParagraph As New Paragraph
            If CompletionItemWrapper.Documentation IsNot Nothing Then
                Dim name = CompletionItemWrapper.CompletionItem.MemberInfo.Name

                If CompilerService.DummyCompiler.TypeInfoBag.Types.TryGetValue(name.ToLowerInvariant(), value) Then
                    For Each item In value.Properties.OrderBy(Function(a) a.Key)
                        detailsParagraph.Inlines.Add(New Image With {
                            .Source = GetIconForSymbolType(LanguageService.SymbolType.Property),
                            .Width = 16.0,
                            .Height = 16.0,
                            .Margin = New Thickness(2.0),
                            .VerticalAlignment = VerticalAlignment.Center
                        })
                        detailsParagraph.Inlines.Add(New TextRun With {
                            .Text = item.Value.Name & vbLf,
                            .Style = _membersStyle
                        })
                    Next

                    For Each item2 In value.Methods.OrderBy(Function(a) a.Key)
                        detailsParagraph.Inlines.Add(New Image With {
                            .Source = GetIconForSymbolType(LanguageService.SymbolType.Method),
                            .Width = 16.0,
                            .Height = 16.0,
                            .Margin = New Thickness(2.0),
                            .VerticalAlignment = VerticalAlignment.Center
                        })
                        detailsParagraph.Inlines.Add(New TextRun With {
                            .Text = item2.Value.Name & vbLf,
                            .Style = _membersStyle
                        })
                    Next

                    For Each item3 In value.Events.OrderBy(Function(a) a.Key)
                        detailsParagraph.Inlines.Add(New Image With {
                            .Source = GetIconForSymbolType(LanguageService.SymbolType.Event),
                            .Width = 16.0,
                            .Height = 16.0,
                            .Margin = New Thickness(2.0),
                            .VerticalAlignment = VerticalAlignment.Center
                        })
                        detailsParagraph.Inlines.Add(New TextRun With {
                            .Text = item3.Value.Name & vbLf,
                            .Style = _membersStyle
                        })
                    Next
                End If
            End If
        End Sub

        Private Sub FillExample()
            Dim detailsParagraph As New Paragraph
            If CompletionItemWrapper.Documentation IsNot Nothing AndAlso Not Equals(CompletionItemWrapper.Documentation.Example, Nothing) Then
                detailsParagraph.Inlines.Add(New TextRun With {
                    .Text = ResourceHelper.GetString("Example"),
                    .Style = _paramsNameStyle
                })

                detailsParagraph.Inlines.Add(New LineBreak())
                '__ = CompletionItemWrapper.Documentation.Example
                detailsParagraph.Inlines.Add(New TextRun With {
                    .Text = CompletionItemWrapper.Documentation.Example,
                    .Style = _codeExampleStyle
                })

                detailsParagraph.Inlines.Add(New LineBreak())
            End If
        End Sub

        Private Sub Notify([property] As String)
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs([property]))
        End Sub

        Private Sub gotoDefinintion_RequestNavigate(sender As Object, e As RequestNavigateEventArgs) Handles gotoDefinintion.RequestNavigate
            MainWindow.popHelp.IsOpen = False
            Dim item = _itemWrapper.CompletionItem
            Dim identifier = item.DefinitionIdintifier
            If identifier.Line = -1 Then ' Select the control on the form
                MainWindow.SelectControl(item.DisplayName)
            Else
                Dim doc = MainWindow.ActiveDocument
                Dim editor = doc.EditorControl
                Dim snapshot = editor.TextView.TextSnapshot
                Dim lineStart = snapshot.GetLineFromLineNumber(identifier.Line).Start
                Dim pos = lineStart + identifier.Column
                doc.EnsureAtTop(pos)
                editor.EditorOperations.SelectCurrentWord()
            End If
        End Sub
    End Class
End Namespace
