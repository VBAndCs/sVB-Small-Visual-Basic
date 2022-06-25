Imports System.Collections.Generic
Imports System.ComponentModel.Composition
Imports System.Text
Imports System.Windows
Imports System.Windows.Media
Imports Microsoft.Nautilus.Core.Undo
Imports Microsoft.Nautilus.Text.Classification
Imports Microsoft.Nautilus.Text.Editor

Namespace Microsoft.Nautilus.Text.Operations
    Public NotInheritable Class EditorCommands

        <Import>
        Public Property EditorOperationsProvider As IEditorOperationsProvider

        <Import>
        Public Property UndoHistoryRegistry As IUndoHistoryRegistry

        <Import>
        Public Property ClassifierAggregatorProvider As IClassifierAggregatorProvider

        <Import>
        Public Property ClassificationFormatMap As IClassificationFormatMap

        Private Function GetUndoHistory(textBuffer As ITextBuffer) As UndoHistory
            Return _undoHistoryRegistry.GetHistory(textBuffer)
        End Function

        <Export(GetType(CommandHandler))>
        <ExportProperty("CommandName", "Microsoft.Editor.Yank.CurrentLine")>
        Public Sub YankCurrentLineHandler(textView As IAvalonTextView)
            Dim editorOperations As IEditorOperations = _editorOperationsProvider.GetEditorOperations(textView)
            Dim undoHistory1 As UndoHistory = GetUndoHistory(textView.TextBuffer)
            editorOperations.YankCurrentLine(undoHistory1)
        End Sub

        <ExportProperty("CommandName", "Microsoft.Editor.Delete.PreviousWord")>
        <Export(GetType(CommandHandler))>
        Public Sub DeletePreviousWordHandler(textView As IAvalonTextView)
            Dim editorOperations As IEditorOperations = _editorOperationsProvider.GetEditorOperations(textView)
            Dim undoHistory1 As UndoHistory = GetUndoHistory(textView.TextBuffer)
            editorOperations.DeleteWordToLeft(undoHistory1)
        End Sub

        <Export(GetType(CommandHandler))>
        <ExportProperty("CommandName", "Microsoft.Editor.Delete.NextWord")>
        Public Sub DeleteNextWordHandler(textView As IAvalonTextView)
            Dim editorOperations As IEditorOperations = _editorOperationsProvider.GetEditorOperations(textView)
            Dim undoHistory1 As UndoHistory = GetUndoHistory(textView.TextBuffer)
            editorOperations.DeleteWordToRight(undoHistory1)
        End Sub

        <Export(GetType(CommandHandler))>
        <ExportProperty("CommandName", "Common.Edit.SelectAll")>
        Public Sub SelectAllHandler(textView As IAvalonTextView)
            Dim editorOperations As IEditorOperations = _editorOperationsProvider.GetEditorOperations(textView)
            editorOperations.SelectAll()
        End Sub

        <Export(GetType(CommandHandler))>
        <ExportProperty("CommandName", "Microsoft.Editor.Select.CurrentWord")>
        Public Sub SelectCurrentWordHandler(textView As IAvalonTextView)
            Dim editorOperations As IEditorOperations = _editorOperationsProvider.GetEditorOperations(textView)
            editorOperations.SelectCurrentWord()
        End Sub

        <ExportProperty("CommandName", "Microsoft.Editor.Move.ToNextWord")>
        <Export(GetType(CommandHandler))>
        Public Sub MoveToNextWordHandler(textView As IAvalonTextView)
            Dim editorOperations As IEditorOperations = _editorOperationsProvider.GetEditorOperations(textView)
            editorOperations.MoveToNextWord([select]:=False)
        End Sub

        <ExportProperty("CommandName", "Microsoft.Editor.Select.ToNextWord")>
        <Export(GetType(CommandHandler))>
        Public Sub SelectToNextWordHandler(textView As IAvalonTextView)
            Dim editorOperations As IEditorOperations = _editorOperationsProvider.GetEditorOperations(textView)
            editorOperations.MoveToNextWord([select]:=True)
        End Sub

        <ExportProperty("CommandName", "Microsoft.Editor.Move.ToPreviousWord")>
        <Export(GetType(CommandHandler))>
        Public Sub MoveToPreviousWordHandler(textView As IAvalonTextView)
            Dim editorOperations As IEditorOperations = _editorOperationsProvider.GetEditorOperations(textView)
            editorOperations.MoveToPreviousWord([select]:=False)
        End Sub

        <Export(GetType(CommandHandler))>
        <ExportProperty("CommandName", "Microsoft.Editor.Select.ToPreviousWord")>
        Public Sub SelectToPreviousWordHandler(textView As IAvalonTextView)
            Dim editorOperations As IEditorOperations = _editorOperationsProvider.GetEditorOperations(textView)
            editorOperations.MoveToPreviousWord([select]:=True)
        End Sub

        <ExportProperty("CommandName", "Microsoft.Editor.Move.ToNextCharacter")>
        <Export(GetType(CommandHandler))>
        Public Sub MoveToNextCharacterHandler(textView As IAvalonTextView)
            Dim editorOperations As IEditorOperations = _editorOperationsProvider.GetEditorOperations(textView)
            editorOperations.MoveCharacterRight([select]:=False)
        End Sub

        <ExportProperty("CommandName", "Microsoft.Editor.Select.ToNextCharacter")>
        <Export(GetType(CommandHandler))>
        Public Sub SelectToNextCharacterHandler(textView As IAvalonTextView)
            Dim editorOperations As IEditorOperations = _editorOperationsProvider.GetEditorOperations(textView)
            editorOperations.MoveCharacterRight([select]:=True)
        End Sub

        <ExportProperty("CommandName", "Microsoft.Editor.Move.ToPreviousCharacter")>
        <Export(GetType(CommandHandler))>
        Public Sub MoveToPreviousCharacterHandler(textView As IAvalonTextView)
            Dim editorOperations As IEditorOperations = _editorOperationsProvider.GetEditorOperations(textView)
            editorOperations.MoveCharacterLeft([select]:=False)
        End Sub

        <Export(GetType(CommandHandler))>
        <ExportProperty("CommandName", "Microsoft.Editor.Select.ToPreviousCharacter")>
        Public Sub SelectToPreviousCharacterHandler(textView As IAvalonTextView)
            Dim editorOperations As IEditorOperations = _editorOperationsProvider.GetEditorOperations(textView)
            editorOperations.MoveCharacterLeft([select]:=True)
        End Sub

        <Export(GetType(CommandHandler))>
        <ExportProperty("CommandName", "Microsoft.Editor.Move.ToNextLine")>
        Public Sub MoveToNextLineHandler(textView As IAvalonTextView)
            Dim editorOperations As IEditorOperations = _editorOperationsProvider.GetEditorOperations(textView)
            editorOperations.MoveLineDown([select]:=False)
        End Sub

        <ExportProperty("CommandName", "Microsoft.Editor.Select.ToNextLine")>
        <Export(GetType(CommandHandler))>
        Public Sub SelectToNextLineHandler(textView As IAvalonTextView)
            Dim editorOperations As IEditorOperations = _editorOperationsProvider.GetEditorOperations(textView)
            editorOperations.MoveLineDown([select]:=True)
        End Sub

        <Export(GetType(CommandHandler))>
        <ExportProperty("CommandName", "Microsoft.Editor.Move.ToPreviousLine")>
        Public Sub MoveToPreviousLineHandler(textView As IAvalonTextView)
            Dim editorOperations As IEditorOperations = _editorOperationsProvider.GetEditorOperations(textView)
            editorOperations.MoveLineUp([select]:=False)
        End Sub

        <ExportProperty("CommandName", "Microsoft.Editor.Select.ToPreviousLine")>
        <Export(GetType(CommandHandler))>
        Public Sub SelectToPreviousLineHandler(textView As IAvalonTextView)
            Dim editorOperations As IEditorOperations = _editorOperationsProvider.GetEditorOperations(textView)
            editorOperations.MoveLineUp([select]:=True)
        End Sub

        <Export(GetType(CommandHandler))>
        <ExportProperty("CommandName", "Microsoft.Editor.Move.ToEndOfLine")>
        Public Sub MoveToEndOfLineHandler(textView As IAvalonTextView)
            Dim editorOperations As IEditorOperations = _editorOperationsProvider.GetEditorOperations(textView)
            editorOperations.MoveToEndOfLine([select]:=False)
        End Sub

        <ExportProperty("CommandName", "Microsoft.Editor.Select.ToEndOfLine")>
        <Export(GetType(CommandHandler))>
        Public Sub SelectToEndOfLineHandler(textView As IAvalonTextView)
            Dim editorOperations As IEditorOperations = _editorOperationsProvider.GetEditorOperations(textView)
            editorOperations.MoveToEndOfLine([select]:=True)
        End Sub

        <ExportProperty("CommandName", "Microsoft.Editor.Move.ToStartOfLine")>
        <Export(GetType(CommandHandler))>
        Public Sub MoveToStartOfLineHandler(textView As IAvalonTextView)
            Dim editorOperations As IEditorOperations = _editorOperationsProvider.GetEditorOperations(textView)
            editorOperations.MoveToStartOfLine([select]:=False)
        End Sub

        <Export(GetType(CommandHandler))>
        <ExportProperty("CommandName", "Microsoft.Editor.Select.ToStartOfLine")>
        Public Sub SelectToStartOfLineHandler(textView As IAvalonTextView)
            Dim editorOperations As IEditorOperations = _editorOperationsProvider.GetEditorOperations(textView)
            editorOperations.MoveToStartOfLine([select]:=True)
        End Sub

        <Export(GetType(CommandHandler))>
        <ExportProperty("CommandName", "Microsoft.Editor.Move.ToStartOfDocument")>
        Public Sub MoveToStartOfDocumentHandler(textView As IAvalonTextView)
            Dim editorOperations As IEditorOperations = _editorOperationsProvider.GetEditorOperations(textView)
            editorOperations.MoveToStartOfDocument([select]:=False)
        End Sub

        <ExportProperty("CommandName", "Microsoft.Editor.Select.ToStartOfDocument")>
        <Export(GetType(CommandHandler))>
        Public Sub SelectToStartOfDocumentHandler(textView As IAvalonTextView)
            Dim editorOperations As IEditorOperations = _editorOperationsProvider.GetEditorOperations(textView)
            editorOperations.MoveToStartOfDocument([select]:=True)
        End Sub

        <Export(GetType(CommandHandler))>
        <ExportProperty("CommandName", "Microsoft.Editor.Move.ToEndOfDocument")>
        Public Sub MoveToEndOfDocumentHandler(textView As IAvalonTextView)
            Dim editorOperations As IEditorOperations = _editorOperationsProvider.GetEditorOperations(textView)
            editorOperations.MoveToEndOfDocument([select]:=False)
        End Sub

        <ExportProperty("CommandName", "Microsoft.Editor.Select.ToEndOfDocument")>
        <Export(GetType(CommandHandler))>
        Public Sub SelectToEndOfDocumentHandler(textView As IAvalonTextView)
            Dim editorOperations As IEditorOperations = _editorOperationsProvider.GetEditorOperations(textView)
            editorOperations.MoveToEndOfDocument([select]:=True)
        End Sub

        <ExportProperty("CommandName", "Microsoft.Editor.Move.CurrentLineToTopOfView")>
        <Export(GetType(CommandHandler))>
        Public Sub MoveCurrentLineToTopOfViewHandler(textView As IAvalonTextView)
            Dim editorOperations As IEditorOperations = _editorOperationsProvider.GetEditorOperations(textView)
            editorOperations.MoveCurrentLineToTop()
        End Sub

        <ExportProperty("CommandName", "Microsoft.Editor.Move.CurrentLineToBottomOfView")>
        <Export(GetType(CommandHandler))>
        Public Sub MoveCurrentLineToBottomOfViewHandler(textView As IAvalonTextView)
            Dim editorOperations As IEditorOperations = _editorOperationsProvider.GetEditorOperations(textView)
            editorOperations.MoveCurrentLineToBottom()
        End Sub

        <Export(GetType(CommandHandler))>
        <ExportProperty("CommandName", "Microsoft.Editor.Select.Clear")>
        Public Sub ClearSelectionHandler(textView As IAvalonTextView)
            Dim editorOperations As IEditorOperations = _editorOperationsProvider.GetEditorOperations(textView)
            editorOperations.ResetSelection()
        End Sub

        <ExportProperty("CommandName", "Common.Edit.Delete")>
        <Export(GetType(CommandHandler))>
        Public Sub DeleteNextCharacterHandler(textView As IAvalonTextView)
            Dim editorOperations As IEditorOperations = _editorOperationsProvider.GetEditorOperations(textView)
            Dim undoHistory1 As UndoHistory = GetUndoHistory(textView.TextBuffer)
            editorOperations.DeleteCharacterToRight(undoHistory1)
        End Sub

        <ExportProperty("CommandName", "Microsoft.Editor.Delete.PreviousCharacter")>
        <Export(GetType(CommandHandler))>
        Public Sub DeletePreviousCharacterHandler(textView As IAvalonTextView)
            Dim editorOperations As IEditorOperations = _editorOperationsProvider.GetEditorOperations(textView)
            Dim undoHistory1 As UndoHistory = GetUndoHistory(textView.TextBuffer)
            editorOperations.DeleteCharacterToLeft(undoHistory1)
        End Sub

        <ExportProperty("CommandName", "Microsoft.Editor.Mode.Overwrite")>
        <Export(GetType(CommandHandler))>
        Public Sub OverwriteModeHandler(textView As IAvalonTextView)
            Dim editorOperations As IEditorOperations = _editorOperationsProvider.GetEditorOperations(textView)
            editorOperations.OverwriteMode = Not editorOperations.OverwriteMode
        End Sub

        <Export(GetType(CommandHandler))>
        <ExportProperty("CommandName", "Microsoft.Editor.Insert.Newline")>
        Public Sub InsertNewlineHandler(textView As IAvalonTextView)
            Dim editorOperations As IEditorOperations = _editorOperationsProvider.GetEditorOperations(textView)
            Dim undoHistory1 As UndoHistory = GetUndoHistory(textView.TextBuffer)
            editorOperations.InsertNewline(undoHistory1)
        End Sub

        <Export(GetType(CommandHandler))>
        <ExportProperty("CommandName", "Microsoft.Editor.Tab.Insert")>
        Public Sub InsertTabHandler(textView As IAvalonTextView)
            Dim editorOperations As IEditorOperations = _editorOperationsProvider.GetEditorOperations(textView)
            Dim undoHistory1 As UndoHistory = GetUndoHistory(textView.TextBuffer)
            If textView.Selection.IsEmpty Then
                editorOperations.InsertTab(undoHistory1)
            Else
                editorOperations.IndentSelection(undoHistory1)
            End If
        End Sub

        <ExportProperty("CommandName", "Microsoft.Editor.Tab.RemovePrevious")>
        <Export(GetType(CommandHandler))>
        Public Sub RemovePreviousTabHandler(textView As IAvalonTextView)
            Dim editorOperations As IEditorOperations = _editorOperationsProvider.GetEditorOperations(textView)
            Dim undoHistory1 As UndoHistory = GetUndoHistory(textView.TextBuffer)
            If textView.Selection.IsEmpty Then
                editorOperations.RemovePreviousTab(undoHistory1)
            Else
                editorOperations.UnindentSelection(undoHistory1)
            End If
        End Sub

        <ExportProperty("CommandName", "Microsoft.Editor.Move.PageUp")>
        <Export(GetType(CommandHandler))>
        Public Sub PageUpHandler(textView As IAvalonTextView)
            Dim editorOperations As IEditorOperations = _editorOperationsProvider.GetEditorOperations(textView)
            editorOperations.PageUp([select]:=False)
        End Sub

        <Export(GetType(CommandHandler))>
        <ExportProperty("CommandName", "Microsoft.Editor.Select.PageUp")>
        Public Sub SelectPageUpHandler(textView As IAvalonTextView)
            Dim editorOperations As IEditorOperations = _editorOperationsProvider.GetEditorOperations(textView)
            editorOperations.PageUp([select]:=True)
        End Sub

        <Export(GetType(CommandHandler))>
        <ExportProperty("CommandName", "Microsoft.Editor.Move.PageDown")>
        Public Sub PageDownHandler(textView As IAvalonTextView)
            Dim editorOperations As IEditorOperations = _editorOperationsProvider.GetEditorOperations(textView)
            editorOperations.PageDown([select]:=False)
        End Sub

        <Export(GetType(CommandHandler))>
        <ExportProperty("CommandName", "Microsoft.Editor.Select.PageDown")>
        Public Sub SelectPageDownHandler(textView As IAvalonTextView)
            Dim editorOperations As IEditorOperations = _editorOperationsProvider.GetEditorOperations(textView)
            editorOperations.PageDown([select]:=True)
        End Sub

        <ExportProperty("CommandName", "Microsoft.Editor.Move.ScrollUp")>
        <Export(GetType(CommandHandler))>
        Public Sub ScrollUpHandler(textView As IAvalonTextView)
            Dim editorOperations As IEditorOperations = _editorOperationsProvider.GetEditorOperations(textView)
            editorOperations.ScrollUpAndMoveCaretIfNecessary()
        End Sub

        <Export(GetType(CommandHandler))>
        <ExportProperty("CommandName", "Microsoft.Editor.Move.ScrollDown")>
        Public Sub ScrollDownHandler(textView As IAvalonTextView)
            Dim editorOperations As IEditorOperations = _editorOperationsProvider.GetEditorOperations(textView)
            editorOperations.ScrollDownAndMoveCaretIfNecessary()
        End Sub

        <ExportProperty("CommandName", "Microsoft.Editor.Transpose.Character")>
        <Export(GetType(CommandHandler))>
        Public Sub TransposeCharacterHandler(textView As IAvalonTextView)
            Dim editorOperations As IEditorOperations = _editorOperationsProvider.GetEditorOperations(textView)
            Dim undoHistory1 As UndoHistory = GetUndoHistory(textView.TextBuffer)
            editorOperations.TransposeCharacter(undoHistory1)
        End Sub

        <Export(GetType(CommandHandler))>
        <ExportProperty("CommandName", "Microsoft.Editor.Transpose.Line")>
        Public Sub TransposeLineHandler(textView As IAvalonTextView)
            Dim editorOperations As IEditorOperations = _editorOperationsProvider.GetEditorOperations(textView)
            Dim undoHistory1 As UndoHistory = GetUndoHistory(textView.TextBuffer)
            editorOperations.TransposeLine(undoHistory1)
        End Sub

        <ExportProperty("CommandName", "Microsoft.Editor.CharacterCase.Lowercase")>
        <Export(GetType(CommandHandler))>
        Public Sub MakeLowercaseHandler(textView As IAvalonTextView)
            Dim editorOperations As IEditorOperations = _editorOperationsProvider.GetEditorOperations(textView)
            Dim undoHistory1 As UndoHistory = GetUndoHistory(textView.TextBuffer)
            editorOperations.MakeLowercase(undoHistory1)
        End Sub

        <Export(GetType(CommandHandler))>
        <ExportProperty("CommandName", "Microsoft.Editor.CharacterCase.Uppercase")>
        Public Sub MakeUppercaseHandler(textView As IAvalonTextView)
            Dim editorOperations As IEditorOperations = _editorOperationsProvider.GetEditorOperations(textView)
            Dim undoHistory1 As UndoHistory = GetUndoHistory(textView.TextBuffer)
            editorOperations.MakeUppercase(undoHistory1)
        End Sub

        <ExportProperty("CommandName", "Common.Edit.Cut")>
        <Export(GetType(CommandHandler))>
        Public Sub CutHandler(textView As IAvalonTextView)
            CopyHandler(textView)
            DeleteNextCharacterHandler(textView)
        End Sub

        <Export(GetType(CommandHandler))>
        <ExportProperty("CommandName", "Common.Edit.Paste")>
        Public Sub PasteHandler(textView As IAvalonTextView)
            Dim editorOperations As IEditorOperations = _editorOperationsProvider.GetEditorOperations(textView)
            Dim undoHistory1 As UndoHistory = GetUndoHistory(textView.TextBuffer)
            editorOperations.Paste(undoHistory1)
        End Sub

        <ExportProperty("CommandName", "Common.Edit.Copy")>
        <Export(GetType(CommandHandler))>
        Public Sub CopyHandler(textView As IAvalonTextView)
            Dim textSnapshot1 As ITextSnapshot = textView.TextSnapshot
            Dim activeSnapshotSpan1 As SnapshotSpan = textView.Selection.ActiveSnapshotSpan
            Dim classifier As IClassifier = ClassifierAggregatorProvider.GetClassifier(textView.TextBuffer)
            Dim classificationSpans As IList(Of ClassificationSpan) = classifier.GetClassificationSpans(activeSnapshotSpan1)
            Dim list1 As New List(Of IClassificationType)
            Dim stringBuilder1 As New StringBuilder("{\rtf1{&H7}nsi{&H7}nsicpg\lang1024{vbLf}oproof65001\uc1 \deff0{{ChrW(12)}onttbl{{ChrW(12)}0{ChrW(12)}nil{ChrW(12)}charset0{ChrW(12)}prq1 Consolas;}}")
            Dim stringBuilder2 As New StringBuilder("{ChrW(12)}s24 ")
            Dim num As Integer = activeSnapshotSpan1.Span.Start

            For Each item As ClassificationSpan In classificationSpans
                Dim span1 As SnapshotSpan = item.GetSpan(textSnapshot1)
                If num < span1.Start Then
                    Dim span2 As New Span(num, Math.Min(span1.Start, span1.End) - num)
                    Dim text As String = textSnapshot1.GetText(span2)
                    stringBuilder2.AppendFormat("{vbCr}f0 {0}", text.Replace("\", "\").Replace("{", "\{").Replace("}", "\}"))
                    num = span2.End
                End If

                Dim span3 = span1.Overlap(activeSnapshotSpan1)
                If span3.HasValue Then
                    Dim text2 As String = textSnapshot1.GetText(span3.Value)
                    If Not list1.Contains(item.ClassificationType) Then
                        list1.Add(item.ClassificationType)
                    End If

                    Dim textProperties As TextFormattingRunProperties = ClassificationFormatMap.GetTextProperties(item.ClassificationType)
                    text2 = text2.Replace("\", "\").Replace("{", "\{").Replace("}", "\}")
                    Dim num2 As Integer = list1.IndexOf(item.ClassificationType)
                    Dim text3 As String = ""
                    Dim text4 As String = ""

                    If textProperties.Typeface.Weight = FontWeights.Bold Then
                        text3 = "{{ChrW(&H8)} "
                        text4 = "}"
                    End If

                    Dim text5 As String = ""
                    Dim text6 As String = ""

                    If textProperties.Typeface.Style = FontStyles.Italic Then
                        text5 = "{\i "
                        text6 = "}"
                    End If

                    stringBuilder2.AppendFormat("{vbCr}f{0} {1}{2}{3}{4}{5}", num2 + 1, text3, text5, text2, text6, text4)
                    num = span1.End
                End If

                If num < span1.End Then
                    Dim span4 As New Span(num, span1.End - num)
                    Dim text7 As String = textSnapshot1.GetText(span4)
                    stringBuilder2.AppendFormat("{vbCr}f0 {0}", text7.Replace("\", "\").Replace("{", "\{").Replace("}", "\}"))
                End If
            Next

            stringBuilder2.Replace(Microsoft.VisualBasic.vbNewLine, "\par ")
            Dim stringBuilder3 As New StringBuilder("{{vbCr}olortbl;")
            For Each item2 As IClassificationType In list1
                Dim textProperties2 As TextFormattingRunProperties = ClassificationFormatMap.GetTextProperties(item2)
                Dim color1 As Color = CType(textProperties2.ForegroundBrush, SolidColorBrush).Color
                stringBuilder3.AppendFormat("\red{0}\green{1}{ChrW(&H8)}lue{2};", color1.R, color1.G, color1.B)
            Next

            stringBuilder1.Append(stringBuilder3.ToString() & "}" & vbCrLf)
            stringBuilder1.Append(stringBuilder2.ToString())
            stringBuilder1.Append("}")
            Dim dataObject1 As New DataObject
            dataObject1.SetData(DataFormats.Rtf, stringBuilder1.ToString())
            dataObject1.SetData(DataFormats.Text, activeSnapshotSpan1.GetText())
            Clipboard.SetDataObject(dataObject1)
        End Sub
    End Class
End Namespace
