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
        Public Shared Current As EditorCommands

        Public Sub New()
            Current = Me
        End Sub

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
            Dim editorOperations = _EditorOperationsProvider.GetEditorOperations(textView)
            Dim undoHistory = GetUndoHistory(textView.TextBuffer)
            editorOperations.YankCurrentLine(undoHistory)
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
            Dim editorOperations = _EditorOperationsProvider.GetEditorOperations(textView)
            Dim undoHistory = GetUndoHistory(textView.TextBuffer)
            editorOperations.DeleteCharacterToRight(undoHistory)
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


        Dim CopyFailed As Boolean

        <ExportProperty("CommandName", "Common.Edit.Cut")>
        <Export(GetType(CommandHandler))>
        Public Sub CutHandler(textView As IAvalonTextView)
            CopyFailed = False
            CopyHandler(textView)
            If Not CopyFailed Then
                DeleteNextCharacterHandler(textView)
            End If
        End Sub

        <Export(GetType(CommandHandler))>
        <ExportProperty("CommandName", "Common.Edit.Paste")>
        Public Sub PasteHandler(textView As IAvalonTextView)
            Dim editorOperations = _EditorOperationsProvider.GetEditorOperations(textView)
            Dim undoHistory = GetUndoHistory(textView.TextBuffer)
            editorOperations.Paste(undoHistory)
        End Sub

        <ExportProperty("CommandName", "Common.Edit.Copy")>
        <Export(GetType(CommandHandler))>
        Public Sub CopyHandler(textView As IAvalonTextView)
            Dim dataObj As DataObject = Nothing

            Try
                Dim snapshot = textView.TextSnapshot
                Dim activeSnapshotSpan = textView.Selection.ActiveSnapshotSpan
                Dim classifier = ClassifierAggregatorProvider.GetClassifier(textView.TextBuffer)
                Dim classificationSpans = classifier.GetClassificationSpans(activeSnapshotSpan)
                Dim classificationTypes As New List(Of IClassificationType)
                Dim sbText As New StringBuilder("{\rtf1\ansi\ansicpg\lang1024\noproof65001\uc1 \deff0{\fonttbl{\f0\fnil\fcharset0\fprq1 Consolas;}}")
                Dim sbClassification As New StringBuilder("\fs24 ")
                Dim num = activeSnapshotSpan.Span.Start

                For Each classiSpan In classificationSpans
                    Dim snapshotSpan = classiSpan.GetSpan(snapshot)
                    If num < snapshotSpan.Start Then
                        Dim span As New Span(num, Math.Min(snapshotSpan.Start, snapshotSpan.End) - num)
                        Dim text = snapshot.GetText(span)
                        sbClassification.AppendFormat("\cf0 {0}", text.Replace("\", "\\").Replace("{", "\{").Replace("}", "\}"))
                        num = span.End
                    End If

                    Dim span3 = snapshotSpan.Overlap(activeSnapshotSpan)
                    If span3.HasValue Then
                        Dim text2 As String = snapshot.GetText(span3.Value)
                        If Not classificationTypes.Contains(classiSpan.ClassificationType) Then
                            classificationTypes.Add(classiSpan.ClassificationType)
                        End If

                        Dim textProperties = ClassificationFormatMap.GetTextProperties(classiSpan.ClassificationType)
                        text2 = text2.Replace("\", "\\").Replace("{", "\{").Replace("}", "\}")
                        Dim num2 = classificationTypes.IndexOf(classiSpan.ClassificationType)
                        Dim text3 = ""
                        Dim text4 = ""

                        If textProperties.Typeface.Weight = FontWeights.Bold Then
                            text3 = "{\b "
                            text4 = "}"
                        End If

                        Dim text5 = ""
                        Dim text6 = ""

                        If textProperties.Typeface.Style = FontStyles.Italic Then
                            text5 = "{\i "
                            text6 = "}"
                        End If

                        sbClassification.AppendFormat("\cf{0} {1}{2}{3}{4}{5}", num2 + 1, text3, text5, text2, text6, text4)
                        num = snapshotSpan.End
                    End If

                    If num < snapshotSpan.End Then
                        Dim span4 As New Span(num, snapshotSpan.End - num)
                        Dim text7 = snapshot.GetText(span4)
                        sbClassification.AppendFormat("\cf0 {0}", text7.Replace("\", "\\").Replace("{", "\{").Replace("}", "\}"))
                    End If
                Next

                sbClassification.Replace(vbNewLine, "\par ")
                Dim sbFormat As New StringBuilder("{\colortbl;" & vbCrLf)

                For Each type In classificationTypes
                    Dim props = ClassificationFormatMap.GetTextProperties(type)
                    Dim color = CType(props.ForegroundBrush, SolidColorBrush).Color
                    sbFormat.AppendFormat("\red{0}\green{1}\blue{2};", color.R, color.G, color.B)
                Next

                sbText.Append(sbFormat.ToString() & "}" & vbCrLf)
                sbText.Append(sbClassification.ToString())
                sbText.Append("}")

                dataObj = New DataObject()
                dataObj.SetData(DataFormats.Rtf, sbText.ToString())
                dataObj.SetData(DataFormats.Text, activeSnapshotSpan.GetText())

            Catch ex As Exception
                MsgBox("Operation has failed. Please try again. " & ex.Message)
                CopyFailed = True
                Return
            End Try

            CopyFailed = Not SetData(dataObj)
        End Sub

        Private Function SetData(dataObj As DataObject) As Boolean
            For i = 1 To 5
                Try
                    Clipboard.SetDataObject(dataObj, True)
                    Return True
                Catch
                    ' Try again after 20 ms
                    System.Threading.Thread.Sleep(20)
                End Try
            Next

            ' The 5 trials to write data faild
            MsgBox("Operation has failed. Please try again.")
            Return False
        End Function
    End Class
End Namespace
