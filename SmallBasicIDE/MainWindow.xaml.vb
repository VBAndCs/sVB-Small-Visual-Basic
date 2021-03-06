﻿Imports Microsoft.Nautilus.Text
Imports Microsoft.SmallBasic.com.smallbasic
Imports Microsoft.SmallBasic.Documents
Imports Microsoft.SmallBasic.LanguageService
Imports Microsoft.SmallBasic.Shell
Imports Microsoft.SmallBasic.Utility
Imports Microsoft.Win32
Imports Microsoft.Windows.Controls
Imports System
Imports System.Collections.Generic
Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.Diagnostics
Imports System.Globalization
Imports System.IO
Imports Nf = Microsoft.SmallBasic.Utility.NotificationButtons
Imports System.Reflection
Imports System.Threading
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input
Imports System.Windows.Media
Imports System.Windows.Media.Imaging
Imports System.Windows.Threading
Imports System.Linq
Imports System.ComponentModel.Composition
Imports RegexBuilder
Imports sb = Microsoft.SmallBasic.WinForms

Namespace Microsoft.SmallBasic
    <Export("MainWindow")>
    Public Class MainWindow
        Private documentTrackerField As DocumentTracker = New DocumentTracker()
        Private currentCompletionItem As CompletionItemWrapper
        Private helpUpdateTimer As DispatcherTimer
        Private currentProgramProcess As Process
        Private mdiViews As ObservableCollection(Of MdiView)
        Private lastSearchedText As String = ""
        Public Shared NewCommand As RoutedUICommand = New RoutedUICommand(ResourceHelper.GetString("NewCommand"), ResourceHelper.GetString("NewCommand"), GetType(MainWindow))
        Public Shared OpenCommand As RoutedUICommand = New RoutedUICommand(ResourceHelper.GetString("OpenCommand"), ResourceHelper.GetString("OpenCommand"), GetType(MainWindow))
        Public Shared SaveCommand As RoutedUICommand = New RoutedUICommand(ResourceHelper.GetString("SaveCommand"), ResourceHelper.GetString("SaveCommand"), GetType(MainWindow))
        Public Shared SaveAsCommand As RoutedUICommand = New RoutedUICommand(ResourceHelper.GetString("SaveAsCommand"), ResourceHelper.GetString("SaveAsCommand"), GetType(MainWindow))
        Public Shared CutCommand As RoutedUICommand = New RoutedUICommand(ResourceHelper.GetString("CutCommand"), ResourceHelper.GetString("CutCommand"), GetType(MainWindow))
        Public Shared CopyCommand As RoutedUICommand = New RoutedUICommand(ResourceHelper.GetString("CopyCommand"), ResourceHelper.GetString("CopyCommand"), GetType(MainWindow))
        Public Shared PasteCommand As RoutedUICommand = New RoutedUICommand(ResourceHelper.GetString("PasteCommand"), ResourceHelper.GetString("PasteCommand"), GetType(MainWindow))
        Public Shared FindCommand As RoutedUICommand = New RoutedUICommand(ResourceHelper.GetString("FindCommand"), ResourceHelper.GetString("FindCommand"), GetType(MainWindow))
        Public Shared FindNextCommand As RoutedUICommand = New RoutedUICommand(ResourceHelper.GetString("FindNextCommand"), ResourceHelper.GetString("FindNextCommand"), GetType(MainWindow))
        Public Shared UndoCommand As RoutedUICommand = New RoutedUICommand(ResourceHelper.GetString("UndoCommand"), ResourceHelper.GetString("UndoCommand"), GetType(MainWindow))
        Public Shared RedoCommand As RoutedUICommand = New RoutedUICommand(ResourceHelper.GetString("RedoCommand"), ResourceHelper.GetString("RedoCommand"), GetType(MainWindow))
        Public Shared FormatCommand As RoutedUICommand = New RoutedUICommand(ResourceHelper.GetString("FormatProgramCommand"), ResourceHelper.GetString("FormatProgramCommand"), GetType(MainWindow))
        Public Shared RunCommand As RoutedUICommand = New RoutedUICommand(ResourceHelper.GetString("RunProgramCommand"), ResourceHelper.GetString("RunProgramCommand"), GetType(MainWindow))
        Public Shared EndProgramCommand As RoutedUICommand = New RoutedUICommand(ResourceHelper.GetString("EndProgramCommand"), ResourceHelper.GetString("EndProgramCommand"), GetType(MainWindow))
        Public Shared StepOverCommand As RoutedUICommand = New RoutedUICommand(ResourceHelper.GetString("StepOverCommand"), ResourceHelper.GetString("StepOverCommand"), GetType(MainWindow))
        Public Shared BreakpointCommand As RoutedUICommand = New RoutedUICommand(ResourceHelper.GetString("BreakpointCommand"), ResourceHelper.GetString("BreakpointCommand"), GetType(MainWindow))
        Public Shared DebugCommand As RoutedUICommand = New RoutedUICommand(ResourceHelper.GetString("DebugCommand"), ResourceHelper.GetString("DebugCommand"), GetType(MainWindow))
        Public Shared WebSaveCommand As RoutedUICommand = New RoutedUICommand(ResourceHelper.GetString("PublishProgramCommand"), ResourceHelper.GetString("PublishProgramCommand"), GetType(MainWindow))
        Public Shared WebLoadCommand As RoutedUICommand = New RoutedUICommand(ResourceHelper.GetString("ImportProgramCommand"), ResourceHelper.GetString("ImportProgramCommand"), GetType(MainWindow))
        Public Shared ExportToVisualBasicCommand As RoutedUICommand = New RoutedUICommand(ResourceHelper.GetString("ExportToVisualBasicCommand"), ResourceHelper.GetString("ExportToVisualBasicCommand"), GetType(MainWindow))

        ''' <summary>
        ''' DocumentTracker
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property DocumentTracker As DocumentTracker
            Get
                Return documentTrackerField
            End Get
        End Property

        Public ReadOnly Property ActiveDocument As TextDocument
            Get
                Dim result As TextDocument = Nothing
                Dim selectedItem = Me.viewsControl.SelectedItem

                If selectedItem IsNot Nothing Then
                    result = selectedItem.Document
                End If

                Return result
            End Get
        End Property

        Public Sub New()
            Me.InitializeComponent()

            mdiViews = New ObservableCollection(Of MdiView)()
            Me.viewsControl.ItemsSource = mdiViews
            AddHandler CompilerService.CurrentCompletionItemChanged, AddressOf OnCurrentCompletionItemChanged
            'OnFileNew(Me, Nothing)
            helpUpdateTimer = New DispatcherTimer(TimeSpan.FromMilliseconds(200.0), DispatcherPriority.ApplicationIdle, AddressOf OnHelpUpdate, Dispatcher)
            'ThreadPool.QueueUserWorkItem(New WaitCallback(AddressOf OnCheckVersion))
            Dim commandLineArgs As String() = Environment.GetCommandLineArgs()

            If commandLineArgs.Length <= 1 Then
                Return
            End If

            Dim text = commandLineArgs(1)

            If Not text.StartsWith("/") Then
                If File.Exists(text) Then
                    OpenDocIfNot(text)
                Else
                    Utility.MessageBox.Show(String.Format(ResourceHelper.GetString("FileNotFound"), text), ResourceHelper.GetString("Title"), "", NotificationButtons.OK, NotificationIcon.Information)
                End If
            End If
        End Sub

        Private Sub OnFileNew(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Dim textDocument As New TextDocument(Nothing)
            textDocument.ContentType = "text.smallbasic"
            DocumentTracker.TrackDocument(textDocument)
            Dim mdiView As MdiView = New MdiView()
            mdiView.Document = textDocument
            Dim item = mdiView
            mdiViews.Add(item)
        End Sub

        Private Sub OnFileOpen(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Dim openFileDialog As OpenFileDialog = New OpenFileDialog()
            openFileDialog.Filter = ResourceHelper.GetString("SmallBasicFileFilter") & "|*.sb;*.smallbasic"

            If openFileDialog.ShowDialog() = True Then
                OpenDocIfNot(openFileDialog.FileName)
            End If
        End Sub

        Private Sub CanFileSave(ByVal sender As Object, ByVal e As CanExecuteRoutedEventArgs)
            e.CanExecute = ActiveDocument IsNot Nothing
        End Sub

        Private Sub OnFileSave(ByVal sender As Object, ByVal e As RoutedEventArgs)
            SaveDocument(ActiveDocument)
        End Sub

        Private Sub OnFileSaveAs(ByVal sender As Object, ByVal e As RoutedEventArgs)
            SaveDocumentAs(ActiveDocument)
        End Sub

        Private Sub CanEditCut(ByVal sender As Object, ByVal e As CanExecuteRoutedEventArgs)
            e.CanExecute = ActiveDocument IsNot Nothing AndAlso Not ActiveDocument.EditorControl.TextView.Selection.IsEmpty
        End Sub

        Private Sub OnEditCut(ByVal sender As Object, ByVal e As RoutedEventArgs)
            ActiveDocument.EditorControl.EditorOperations.CutSelection(ActiveDocument.UndoHistory)
        End Sub

        Private Sub CanEditCopy(ByVal sender As Object, ByVal e As CanExecuteRoutedEventArgs)
            e.CanExecute = ActiveDocument IsNot Nothing AndAlso Not ActiveDocument.EditorControl.TextView.Selection.IsEmpty
        End Sub

        Private Sub OnEditCopy(ByVal sender As Object, ByVal e As RoutedEventArgs)
            ActiveDocument.EditorControl.EditorOperations.CopySelection()
        End Sub

        Private Sub CanEditPaste(ByVal sender As Object, ByVal e As CanExecuteRoutedEventArgs)
            e.CanExecute = ActiveDocument IsNot Nothing AndAlso ActiveDocument.EditorControl.EditorOperations.CanPaste
        End Sub

        Private Sub OnEditPaste(ByVal sender As Object, ByVal e As RoutedEventArgs)
            ActiveDocument.EditorControl.EditorOperations.Paste(ActiveDocument.UndoHistory)
        End Sub

        Private Sub CanEditUndo(ByVal sender As Object, ByVal e As CanExecuteRoutedEventArgs)
            e.CanExecute = ActiveDocument IsNot Nothing AndAlso ActiveDocument.UndoHistory.CanUndo
        End Sub

        Private Sub OnEditUndo(ByVal sender As Object, ByVal e As RoutedEventArgs)
            ActiveDocument.UndoHistory.Undo(1)
        End Sub

        Private Sub CanEditRedo(ByVal sender As Object, ByVal e As CanExecuteRoutedEventArgs)
            e.CanExecute = ActiveDocument IsNot Nothing AndAlso ActiveDocument.UndoHistory.CanRedo
        End Sub

        Private Sub OnEditRedo(ByVal sender As Object, ByVal e As RoutedEventArgs)
            ActiveDocument.UndoHistory.Redo(1)
        End Sub

        Private Sub OnCloseItem(sender As Object, e As RequestCloseEventArgs)
            Dim mdiView = TryCast(e.Item, MdiView)
            DoCloseDoc(mdiView)
        End Sub

        Sub DoCloseDoc(mdiView As MdiView)
            If mdiView IsNot Nothing AndAlso CloseDocument(mdiView.Document) Then
                mdiViews.Remove(mdiView)
                If mdiViews.Count > 0 Then
                    If viewsControl.SelectedItem Is Nothing Then
                        viewsControl.ChangeSelection(mdiViews(mdiViews.Count - 1))
                    End If
                End If
            End If

        End Sub

        Private Sub CanExportToVisualBasic(ByVal sender As Object, ByVal e As CanExecuteRoutedEventArgs)
            If ActiveDocument IsNot Nothing AndAlso ActiveDocument.Text.Trim().Length > 0 Then
                e.CanExecute = True
            Else
                e.CanExecute = False
            End If
        End Sub

        Private Sub OnExportToVisualBasic(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Dim exportToVBDialog As ExportToVBDialog = New ExportToVBDialog(ActiveDocument)
            exportToVBDialog.Owner = Me
            exportToVBDialog.ShowDialog()
        End Sub

        Private Sub OnFind(ByVal sender As Object, ByVal e As RoutedEventArgs)
            If ActiveDocument Is Nothing Then
                Return
            End If

            Dim messageBox As Utility.MessageBox = New Utility.MessageBox()
            messageBox.Description = ResourceHelper.GetString("TextToSearch")
            messageBox.Title = ResourceHelper.GetString("FindCommand")
            Dim text = lastSearchedText

            If Not ActiveDocument.EditorControl.TextView.Selection.IsEmpty Then
                Dim text2 As String = ActiveDocument.EditorControl.TextView.Selection.ActiveSnapshotSpan.GetText()

                If Not text2.Contains(VisualBasic.Constants.vbCr) AndAlso Not text2.Contains(VisualBasic.Constants.vbLf) Then
                    text = text2
                End If
            End If

            Dim textBox As New TextBox() With {
                .Text = text,
                .FontSize = 20.0,
                .FontFamily = New FontFamily("Consolas"),
                .Foreground = Brushes.DimGray,
                .Margin = New Thickness(0.0, 4.0, 4.0, 4.0),
                .MinWidth = 300.0
            }
            messageBox.OptionalContent = textBox
            messageBox.NotificationButtons = Nf.Close Or Nf.OK
            messageBox.okButton.Content = ResourceHelper.GetString("FindCommand")
            messageBox.NotificationIcon = NotificationIcon.Custom
            messageBox.iconImageControl.Source = New BitmapImage(New Uri("pack://application:,,/SB;component/Resources/Search.png"))
            textBox.SelectAll()
            textBox.Focus()
            Dim notificationButtons As NotificationButtons = messageBox.Display()

            If notificationButtons = NotificationButtons.OK Then
                lastSearchedText = textBox.Text

                If Not ActiveDocument.EditorControl.HighlightNextMatch(lastSearchedText, ignoreCase:=True) Then
                    Console.Beep()
                End If
            End If
        End Sub

        Private Sub OnFindNext(ByVal sender As Object, ByVal e As RoutedEventArgs)
            If ActiveDocument IsNot Nothing Then
                If String.IsNullOrEmpty(lastSearchedText) Then
                    OnFind(sender, e)
                ElseIf Not ActiveDocument.EditorControl.HighlightNextMatch(lastSearchedText, ignoreCase:=True) Then
                    Console.Beep()
                End If
            End If
        End Sub

        Private Sub OnFormatProgram(ByVal sender As Object, ByVal e As RoutedEventArgs)
            If ActiveDocument IsNot Nothing Then
                CompilerService.FormatDocument(ActiveDocument.TextBuffer)
            End If
        End Sub

        Private Sub CanRunProgram(ByVal sender As Object, ByVal e As CanExecuteRoutedEventArgs)
            e.CanExecute = ActiveDocument IsNot Nothing
        End Sub

        Private Sub OnProgramRun(sender As Object, e As RoutedEventArgs)
            RunProgram()
        End Sub

        Private Sub OnProgramEnd(ByVal sender As Object, ByVal e As RoutedEventArgs)
            If currentProgramProcess IsNot Nothing AndAlso Not currentProgramProcess.HasExited Then
                currentProgramProcess.Kill()
                currentProgramProcess = Nothing
                Me.programRunningOverlay.Visibility = Visibility.Hidden
            End If
        End Sub

        Private Sub OnStepOver(ByVal sender As Object, ByVal e As RoutedEventArgs)
            ActiveDocument.Errors.Clear()
            CompilerService.Compile(ActiveDocument.Text, ActiveDocument.Errors)

            If ActiveDocument.Errors.Count = 0 Then
                Dim debugger = ProgramDebugger.GetDebugger(ActiveDocument)
                debugger.StepOver()
            End If
        End Sub

        Private Sub OnToggleBreakpoint(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Dim debugger = ProgramDebugger.GetDebugger(ActiveDocument)
            Dim stMarkerProvider = StatementMarkerProvider.GetStatementMarkerProvider(ActiveDocument.EditorControl.TextView)
            Dim lineFromPosition = ActiveDocument.TextBuffer.CurrentSnapshot.GetLineFromPosition(ActiveDocument.EditorControl.TextView.Caret.Position.CharacterIndex)

            If debugger.Breakpoints.Contains(lineFromPosition.LineNumber) Then
                debugger.Breakpoints.Remove(lineFromPosition.LineNumber)
                Dim array As StatementMarker() = stMarkerProvider.Markers.ToArray()

                For Each statementMarker In array
                    Dim containingLine As ITextSnapshotLine = statementMarker.Span.GetStartPoint(ActiveDocument.TextBuffer.CurrentSnapshot).GetContainingLine()

                    If containingLine.LineNumber = lineFromPosition.LineNumber Then
                        stMarkerProvider.RemoveMarker(statementMarker)
                    End If
                Next
            Else
                debugger.Breakpoints.Add(lineFromPosition.LineNumber)
                stMarkerProvider.AddStatementMarker(New StatementMarker(New TextSpan(ActiveDocument.TextBuffer.CurrentSnapshot, lineFromPosition.Start, lineFromPosition.Length, SpanTrackingMode.EdgeInclusive), Colors.Red))
            End If
        End Sub

        Private Sub CanWebSave(ByVal sender As Object, ByVal e As CanExecuteRoutedEventArgs)
            e.CanExecute = ActiveDocument IsNot Nothing
        End Sub

        Private Sub OnWebSave(ByVal sender As Object, ByVal e As RoutedEventArgs)
            PublishDocument(ActiveDocument)
        End Sub

        Private Sub OnWebLoad(ByVal sender As Object, ByVal e As RoutedEventArgs)
            ImportDocument()
        End Sub

        Protected Overrides Sub OnClosing(ByVal e As CancelEventArgs)
            For Each item In New List(Of MdiView)(mdiViews)
                If Not CloseDocument(item.Document) Then
                    e.Cancel = True
                    Return
                End If

                mdiViews.Remove(item)
            Next

            tabDesigner.IsSelected = True
            For i = 0 To DiagramHelper.Designer.Pages.Count - 1
                If Not DiagramHelper.Designer.ClosePage(False) Then
                    e.Cancel = True
                    Exit For
                End If
            Next

            MyBase.OnClosing(e)
        End Sub

        Private Function OpenCodeFile(ByVal filePath As String) As TextDocument
            Dim doc As New TextDocument(filePath)
            DocumentTracker.TrackDocument(doc)
            Dim mdiView As New MdiView()
            mdiView.Document = doc
            mdiViews.Add(mdiView)
            formDesigner.CodeFilePath = filePath
            Return doc
        End Function

        Private Function SaveDocument(ByVal document As TextDocument) As Boolean
            If document.IsNew OrElse (document.Form <> "" AndAlso formDesigner.FileName = "") Then
                Return SaveDocumentAs(document)
            End If

            Try
                Dim designIsDirty = document.Form <> "" AndAlso formDesigner.HasChanges
                If designIsDirty Then SaveDesignInfo(document)
                If document.IsDirty Then document.Save()
                Return True
            Catch ex As Exception
                Dim notificationButtons = Utility.MessageBox.Show(ResourceHelper.GetString("SaveFailed"), "Small Basic", String.Format(CultureInfo.CurrentCulture, ResourceHelper.GetString("SaveFailedReason"), New Object(0) {ex.Message}), Nf.No Or Nf.Yes, NotificationIcon.Information)

                If notificationButtons = Nf.Yes Then
                    Return SaveDocumentAs(document)
                End If

                Return False
            End Try
        End Function

        Private Function SaveDocumentAs(document As TextDocument) As Boolean
            Dim saveFileDialog As SaveFileDialog = New SaveFileDialog()
            saveFileDialog.Filter = ResourceHelper.GetString("SmallBasicFileFilter") & "|*.sb"
            saveFileDialog.FileName = document.Form

            If saveFileDialog.ShowDialog() Then
                Try

                    Dim FileName = saveFileDialog.FileName
                    If File.Exists(FileName) Then File.Delete(FileName)
                    document.SaveAs(FileName)

                    If document.OpenedInDesigner Then
                        Dim newFormName = Path.GetFileNameWithoutExtension(saveFileDialog.FileName)
                        Dim newDir = Path.GetDirectoryName(saveFileDialog.FileName)
                        Dim newFilePath = Path.Combine(newDir, newFormName)

                        FileName = newFilePath & ".sb.gen"
                        If File.Exists(FileName) Then File.Delete(FileName)

                        FileName = newFilePath & ".xaml"
                        If File.Exists(FileName) Then File.Delete(FileName)

                        formDesigner.FileName = FileName
                        formDesigner.CodeFilePath = saveFileDialog.FileName
                        formDesigner.HasChanges = True ' Force saving
                        SaveDesignInfo(document)
                    End If

                    Return True
                Catch ex As Exception
                    Dim notificationButtons = Utility.MessageBox.Show(ResourceHelper.GetString("SaveFailed"), ResourceHelper.GetString("Title"), String.Format(CultureInfo.CurrentUICulture, ResourceHelper.GetString("SaveFailedReason"), New Object(0) {ex.Message}), Nf.No Or Nf.Yes, NotificationIcon.Information)

                    If notificationButtons = Nf.Yes Then
                        Return SaveDocumentAs(document)
                    End If

                    Return False
                End Try
            End If

            Return False
        End Function

        Private Function CloseDocument(document As TextDocument) As Boolean
            If document.IsDirty Then
                tabCode.IsSelected = True
                Select Case Utility.MessageBox.Show(ResourceHelper.GetString("SaveDocumentBeforeClosing"), ResourceHelper.GetString("Title"), document.Title & ResourceHelper.GetString("DocumentModified"), Nf.Cancel Or Nf.No Or Nf.Yes, NotificationIcon.Information)
                    Case Nf.Yes
                        Return SaveDocument(document)
                    Case Nf.No
                        Return True
                    Case Else
                        Return False
                End Select
            End If

            Return True
        End Function


        Private Sub RunProgram()
            Try
                Dim doc = ActiveDocument
                Dim code As String
                Dim outputFileName = GetOutputFileName(doc)

                ' We need tp pass a value to the forms module in the class library, 
                ' but the shared fileds will not work,
                ' because the generated exe will crated the module in another domain
                ' sp, we will pass the value via registery
                VisualBasic.Interaction.SaveSetting("sVb", "Designer", "CodeFilePath", outputFileName)

                Dim offset = 0

                Dim gen As String
                If doc.OpenedInDesigner Then
                    If doc.Text.Contains("'@Form Hints:") Then
                        doc.ParseFormHints()
                        gen = doc.GetCodeBehind(True)
                    Else
                        gen = doc.GenerateCodeBehind(formDesigner.Items, False)
                    End If
                Else
                    gen = doc.GetCodeBehind(True)
                End If

                If gen <> "" Then
                    offset = CountLines(gen) + 1
                    code = gen & Environment.NewLine & doc.Text
                Else
                    code = doc.Text
                End If


                doc.Errors.Clear()
                Dim errors As List(Of [Error])

                If Not RunProgram(code, errors, outputFileName) Then
                    For Each err As [Error] In errors
                        If err.Line = -1 Then
                            doc.Errors.Add(err.Description)
                        Else
                            doc.Errors.Add($"{err.Line - offset + 1},{err.Column + 1}: {err.Description}")
                        End If
                    Next
                End If
            Catch ex As Exception
                Utility.MessageBox.Show(ResourceHelper.GetString("FailedToCreateOutputFile"), ResourceHelper.GetString("Title"), String.Format(CultureInfo.CurrentUICulture, ResourceHelper.GetString("FailedToCreateOutputFileReason"), New Object(0) {ex.Message}), NotificationButtons.Close, NotificationIcon.Error)
            Finally
                VisualBasic.Interaction.DeleteSetting("sVb", "Designer", "CodeFilePath")
            End Try
        End Sub

        Private Function RunProgram(ByVal code As String, ByRef errors As List(Of [Error]), outputFileName As String) As Boolean
            Dim doc = ActiveDocument
            errors = Compile(code, outputFileName)
            If errors.Count = 0 Then
                Thread.Sleep(500)
                currentProgramProcess = Process.Start(outputFileName)
                currentProgramProcess.EnableRaisingEvents = True

                AddHandler currentProgramProcess.Exited,
                    Sub()
                        Dispatcher.BeginInvoke(DispatcherPriority.Normal, CType(
                             Function()
                                 Me.programRunningOverlay.Visibility = Visibility.Hidden
                                 currentProgramProcess = Nothing

                                 If doc.IsNew Then
                                     Try
                                         File.Delete(outputFileName)
                                     Catch
                                     End Try
                                 End If

                                 doc.Focus()
                                 Return True
                             End Function,
                             DispatcherOperationCallback), Nothing)
                    End Sub

                Me.processRunningMessage.Text = String.Format(ResourceHelper.GetString("ProgramRunning"), doc.Title)
                Me.programRunningOverlay.Visibility = Visibility.Visible
                Me.endProgramButton.Focus()

            ElseIf doc.Form <> "" Then
                Return PreCompile(code, errors, outputFileName)
            Else
                Return False
            End If

            Return True
        End Function


        Private Sub PublishDocument(ByVal document As TextDocument)
            Try
                Cursor = Cursors.Wait
                Dim service As Service = New Service()
                Dim text = service.SaveProgram("", document.Text, document.BaseId)

                If Equals(text, "error") Then
                    Utility.MessageBox.Show(ResourceHelper.GetString("FailedToPublishToWeb"), ResourceHelper.GetString("Title"), ResourceHelper.GetString("PublishToWebFailedReason"), NotificationButtons.Close, NotificationIcon.Error)
                Else
                    Dim publishProgramDialog As PublishProgramDialog = New PublishProgramDialog(text)
                    publishProgramDialog.Owner = Me
                    publishProgramDialog.ShowDialog()
                End If

            Catch ex As Exception
                Utility.MessageBox.Show(ResourceHelper.GetString("FailedToPublishToWeb"), ResourceHelper.GetString("Title"), String.Format(CultureInfo.CurrentUICulture, ResourceHelper.GetString("ReasonForFailure"), New Object(0) {ex.Message}), NotificationButtons.Close, NotificationIcon.Error)
            Finally
                Cursor = Cursors.Arrow
            End Try
        End Sub

        Private Sub ImportDocument()
            Try
                Dim messageBox As New Utility.MessageBox()
                messageBox.Description = ResourceHelper.GetString("ImportFromWeb")
                messageBox.Title = ResourceHelper.GetString("Title")
                Dim stackPanel As StackPanel = New StackPanel()
                stackPanel.Orientation = Orientation.Vertical
                Dim stackPanel2 = stackPanel
                Dim textBlock As New TextBlock()
                textBlock.Text = ResourceHelper.GetString("ImportLocationOfProgramOnWeb")
                textBlock.Margin = New Thickness(0.0, 0.0, 4.0, 4.0)
                Dim element = textBlock
                Dim textBox As New TextBox() With {
                    .FontSize = 32.0,
                    .FontWeight = FontWeights.Bold,
                    .FontFamily = New FontFamily("Courier New"),
                    .Foreground = Brushes.DimGray,
                    .Margin = New Thickness(0.0, 4.0, 4.0, 4.0),
                    .MinWidth = 300.0
                }
                stackPanel2.Children.Add(element)
                stackPanel2.Children.Add(textBox)
                messageBox.OptionalContent = stackPanel2
                messageBox.NotificationButtons = NotificationButtons.Cancel Or NotificationButtons.OK
                messageBox.NotificationIcon = NotificationIcon.Information
                textBox.Focus()

                If messageBox.Display() = NotificationButtons.OK Then
                    Dim service As Service = New Service()
                    Dim baseId As String = textBox.Text.Trim()
                    Dim code = service.LoadProgram(baseId)

                    If Equals(code, "error") Then
                        Utility.MessageBox.Show(ResourceHelper.GetString("FailedToImportFromWeb"), ResourceHelper.GetString("Title"), ResourceHelper.GetString("ImportFromWebFailedReason"), NotificationButtons.Close, NotificationIcon.Error)
                    Else
                        code = code.Replace(VisualBasic.Constants.vbLf, VisualBasic.Constants.vbCrLf)
                        Dim newDocument As New TextDocument(Nothing)
                        newDocument.ContentType = "text.smallbasic"
                        newDocument.BaseId = baseId
                        newDocument.TextBuffer.Insert(0, code)
                        Dim service2 As Service = New Service()
                        AddHandler service2.GetProgramDetailsCompleted, Sub(ByVal o, ByVal e)
                                                                            Dim result = e.Result
                                                                            result.Category = ResourceHelper.GetString("Category" & result.Category)
                                                                            newDocument.ProgramDetails = result
                                                                        End Sub

                        service2.GetProgramDetailsAsync(baseId)
                        DocumentTracker.TrackDocument(newDocument)
                        Dim mdiView As New MdiView()
                        mdiView.Document = newDocument
                        mdiViews.Add(mdiView)
                    End If
                End If

            Catch ex As Exception
                Utility.MessageBox.Show(ResourceHelper.GetString("FailedToImportFromWeb"), ResourceHelper.GetString("Title"), String.Format(CultureInfo.CurrentUICulture, ResourceHelper.GetString("ReasonForFailure"), New Object(0) {ex.Message}), NotificationButtons.Close, NotificationIcon.Error)
            Finally
                Cursor = Cursors.Arrow
            End Try
        End Sub

        Private Function GetOutputFileName(document As TextDocument) As String
            If document.IsNew Then
                Dim tempFileName As String = Path.GetTempFileName()
                File.Move(tempFileName, tempFileName & ".exe")
                Return tempFileName & ".exe"
            End If

            Dim fileNameWithoutExtension = Path.GetFileNameWithoutExtension(document.FilePath)
            Dim directoryName = Path.GetDirectoryName(document.FilePath)
            Return Path.Combine(directoryName, fileNameWithoutExtension & ".exe")
        End Function

        Private Sub OnCurrentCompletionItemChanged(ByVal sender As Object, ByVal e As CurrentCompletionItemChangedEventArgs)
            currentCompletionItem = e.CurrentCompletionItem

            If currentCompletionItem IsNot Nothing Then
                helpUpdateTimer.Start()
            End If
        End Sub

        Private Sub OnHelpUpdate(ByVal sender As Object, ByVal e As EventArgs)
            helpUpdateTimer.Stop()

            If currentCompletionItem IsNot Nothing Then
                Me.helpPanel.CompletionItemWrapper = currentCompletionItem
            End If
        End Sub

        Private Sub OnCheckVersion(state As Object)
            Thread.Sleep(20000)

            Try
                Dim service As Service = New Service()
                Dim currentVersion As String = service.GetCurrentVersion()
                Dim version As Version = New Version(currentVersion)
                Dim version2 As Version = Assembly.GetExecutingAssembly().GetName().Version

                'If version.CompareTo(version2) > 0 Then
                '    Dispatcher.BeginInvoke(CType(Sub() Me.updateAvailable.Visibility = Visibility.Visible, Action))
                'End If

            Catch
            End Try
        End Sub

        Private Sub OnClickNewVersionAvailable(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Process.Start("http://smallbasic.com/download.aspx")
        End Sub

        Function OpenDocIfNot(FilePath As String) As TextDocument
            Dim docPath = Path.GetFullPath(FilePath)
            For Each view As MdiView In Me.viewsControl.Items
                If view.Document.FilePath = docPath Then
                    viewsControl.ChangeSelection(view)
                    Return view.Document
                End If
            Next

            Return OpenCodeFile(FilePath)
        End Function

        Private Sub tabCode_Selected(sender As Object, e As RoutedEventArgs)
            SaveDesignInfo()

            ' Note this prop isn't changed yet
            tabDesigner.IsSelected = False
            UpdateTitle()
        End Sub

        Private Sub tabDesigner_Selected(sender As Object, e As RoutedEventArgs)
            UpdateTitle()
        End Sub

        Dim _projectPath As String

        Private Function GetProjectPath() As String
            If _projectPath <> "" Then Return _projectPath

            Dim tmpPath = "UnSaved"
            If Not IO.Directory.Exists(tmpPath) Then IO.Directory.CreateDirectory(tmpPath)

            Dim n = 1
            Do
                Dim projectName = "Project" & n
                _projectPath = Path.Combine(tmpPath, projectName)
                If Not Directory.Exists(_projectPath) Then
                    Directory.CreateDirectory(_projectPath)
                    Exit Do
                End If
                n += 1
            Loop
            Return _projectPath
        End Function

        Private Sub SaveDesignInfo(Optional doc As TextDocument = Nothing, Optional openDoc As Boolean = True)
            Dim formName As String
            Dim xamlPath As String
            Dim formPath As String

            OpeningDoc = True

            If openDoc AndAlso formDesigner.CodeFilePath <> "" AndAlso Not formDesigner.HasChanges Then
                doc = OpenDocIfNot(formDesigner.CodeFilePath)
                If doc.Form = "" Then doc.Form = formDesigner.Name
                If doc.PageKey = "" Then doc.PageKey = formDesigner.PageKey
                OpeningDoc = False
                Return
            End If

            If formDesigner.FileName = "" Then
                If formDesigner.CodeFilePath = "" Then
                    Dim projectPath = GetProjectPath()
                    xamlPath = ""
                    formName = formDesigner.Name
                    Do
                        xamlPath = Path.Combine(projectPath, formName)
                        If Not Directory.Exists(xamlPath) Then
                            Directory.CreateDirectory(xamlPath)
                            Exit Do
                        End If
                        formName = DiagramHelper.Designer.GetTempFormName().Replace("KEY", "Form")
                    Loop
                    formPath = Path.Combine(xamlPath, formName)

                Else
                    If doc Is Nothing Then doc = GetDoc(formDesigner.CodeFilePath, openDoc)
                    formName = If(doc.Form, formDesigner.Name)
                    xamlPath = Path.GetDirectoryName(formDesigner.CodeFilePath)
                    formPath = formDesigner.CodeFilePath.Substring(0, formDesigner.CodeFilePath.Length - 3)
                End If

                formDesigner.Name = formName
                formDesigner.DoSave(formPath & ".xaml")
                IO.File.Create(formPath & ".sb").Close()

            Else
                xamlPath = Path.GetDirectoryName(formDesigner.FileName)
                formPath = formDesigner.FileName.Substring(0, formDesigner.FileName.Length - 5)

                If doc Is Nothing Then doc = GetDoc(formPath & ".sb", openDoc)
                formName = If(doc.Form, formDesigner.Name)
                formDesigner.Name = formName

                If formDesigner.HasChanges Then formDesigner.DoSave()

            End If

            If doc Is Nothing Then doc = GetDoc(formPath & ".sb", openDoc)
            doc.Form = formName
            doc.PageKey = formDesigner.PageKey
            formDesigner.CodeFilePath = doc.FilePath

            IO.File.WriteAllText(formPath & ".sb.gen", doc.GenerateCodeBehind(formDesigner.Items, True))
            OpeningDoc = False
        End Sub

        Function GetDoc(codeFilePath As String, openDoc As Boolean) As TextDocument
            If Not File.Exists(codeFilePath) Then File.Create(codeFilePath).Close()
            Return If(openDoc,
                    OpenDocIfNot(codeFilePath),
                    New TextDocument(codeFilePath))
        End Function

        Private Sub MainWindow_Closed(sender As Object, e As EventArgs) Handles Me.Closed
            For Each d In IO.Directory.GetDirectories("UnSaved")
                Try
                    Global.My.Computer.FileSystem.DeleteDirectory(d, VisualBasic.FileIO.DeleteDirectoryOption.DeleteAllContents)
                Catch
                End Try
            Next
        End Sub


        Private ReadOnly WordRgex As New Verex(Patterns.Symbols.AnyWord)
        Private ReadOnly OpenBracketRegex As New Verex(Patterns.NoneOrMany(Patterns.Symbols.WhiteSpace) + "(")
        Private ReadOnly MethodRegex As New Verex(Patterns.Symbols.AnyWord + Patterns.NoneOrMany(Patterns.Symbols.WhiteSpace) + "(")

        Private Function PreCompile(code As String, ByRef errors As List(Of [Error]), outputFileName As String) As Boolean
            Dim ReRun = False
            Dim lines = New List(Of String)(code.Split(New String(0) {Environment.NewLine}, StringSplitOptions.None))
            Dim doc = ActiveDocument
            Dim num = errors.Count - 1
            Dim i As Integer = num
            For i = errors.Count - 1 To 0 Step -1
                Dim err = errors(i)
                Dim errMsg = err.Description
                If Not errMsg.StartsWith("Cannot find object", StringComparison.InvariantCultureIgnoreCase) Then Continue For

                Dim lineNum = err.Line
                Dim charNum = err.Column
                Dim pos1 = errMsg.LastIndexOf("'", errMsg.Length - 3) + 1
                Dim obj As String = errMsg.Substring(pos1, errMsg.Length - pos1 - 2).ToLower()
                If Not doc.ControlsInfo.ContainsKey(obj) Then Continue For

                Dim line = lines(lineNum)

                If line.Substring(charNum, obj.Length + 1).ToLower() = obj + "." Then
                    Dim prevText = If(charNum = 0, "", line.Substring(0, charNum))
                    Dim methodPos = charNum + obj.Length + 1
                    Dim nextText = line.Substring(methodPos)

                    Dim match = MethodRegex.Match(nextText)
                    If match.Success AndAlso match.Index = 0 Then ' Method Call
                        Dim method = nextText.Substring(0, WordRgex.Match(nextText).Length)
                        Dim contents = GetBalancedBrackets(nextText)
                        If contents Is Nothing OrElse contents.Count = 0 Then
                            errors(i) = New [Error](err.Line, methodPos, "Wrong brackets pairs")
                            Continue For
                        End If

                        Dim params = contents(0).Value.Trim(" "c, Convert.ToChar(8))
                        Dim RestText = If(contents(0).Index + contents(0).Length > nextText.Length - 2, "",
                                        nextText.Substring(contents(0).Index + contents(0).Length + 1))

                        Dim methodInfo = sb.PreCompiler.GetMethodInfo(doc.ControlsInfo(obj), method)
                        Dim ModuleName = methodInfo.Module
                        If ModuleName = "" Then
                            errors(i) = New [Error](err.Line, methodPos, $"`{method}` doesn't exist.")
                            Continue For
                        End If

                        Dim paramsCount = If(params = "", 0, params.Length - params.Replace(",", "").Length + 1)
                        If methodInfo.ParamsCount = 0 OrElse methodInfo.ParamsCount <= paramsCount Then
                            errors(i) = New [Error](err.Line, methodPos, "Wrong number of parameters.")
                            Continue For
                        End If

                        Select Case methodInfo.ParamsCount
                            Case 1
                                lines(lineNum) = prevText &
                                        $"{ModuleName}.{method}({obj})" &
                                        RestText
                            Case paramsCount + 1
                                lines(lineNum) = prevText &
                                            $"{ModuleName}.{method}({obj}, {params})" &
                                            RestText
                            Case 2
                                lines(lineNum) = prevText &
                                            $"{ModuleName}.{method}({doc.Form}, {obj})" &
                                            RestText
                            Case Else
                                lines(lineNum) = prevText &
                                            $"{ModuleName}.{method}({doc.Form}, {obj}, {params})" &
                                            RestText
                        End Select

                        errors.RemoveAt(i)
                        ReRun = True

                    ElseIf prevText.Trim(" "c, Convert.ToChar(8)) = "" Then 'Property Set
                        pos1 = nextText.IndexOf("="c)
                        If pos1 = -1 Then
                            errors(i) = New [Error](err.Line, methodPos, $"Expected `=` and a value to set the property")
                            Continue For
                        End If

                        Dim result = WordRgex.Match(nextText)
                        If Not result.Success OrElse result.Index > 0 Then
                            errors(i) = New [Error](err.Line, methodPos, $"Expected a property name.")
                            Continue For
                        End If

                        Dim L = pos1 - result.Length

                        If L = 0 OrElse nextText.Substring(result.Length, L).Trim(" "c, Convert.ToChar(8)) = "" Then
                            Dim eventInfo = sb.PreCompiler.GetMethodInfo(doc.ControlsInfo(obj), result.Value)
                            If eventInfo.Module = "" Then
                                Dim method = $"Set{result.Value}"
                                Dim methodInfo = sb.PreCompiler.GetMethodInfo(doc.ControlsInfo(obj), method)
                                Dim ModuleName = methodInfo.Module
                                If ModuleName = "" Then
                                    errors(i) = New [Error](err.Line, methodPos, $"`{result.Value}` doesn't exist.")
                                    Continue For
                                End If

                                Select Case methodInfo.ParamsCount
                                    Case 2
                                        lines(lineNum) = prevText &
                                            $"{ModuleName}.{method}({obj}, {nextText.Substring(pos1 + 1).Trim})"
                                    Case 3
                                        lines(lineNum) = prevText &
                                             $"{ModuleName}.{method}({doc.Form}, {obj}, {nextText.Substring(pos1 + 1).Trim})"
                                    Case Else
                                        errors(i) = New [Error](err.Line, methodPos, $"`{method}` definition is not supported.")
                                        Continue For
                                End Select

                                errors.RemoveAt(i)
                                ReRun = True
                            Else ' Event
                                Dim ModuleName = eventInfo.Module
                                lines.Insert(lineNum, $"Control.HandleEvents({doc.Form}, {obj})")
                                lines(lineNum + 1) = prevText & $"{ModuleName}.{nextText}"
                                errors.RemoveAt(i)
                                ReRun = True
                            End If
                        End If

                    Else 'Property Get          
                        match = WordRgex.Match(nextText)
                        If Not match.Success OrElse match.Index > 0 Then
                            errors(i) = New [Error](err.Line, methodPos, "property name is expected.")
                            Continue For
                        End If

                        Dim method = $"Get{match.Value}"
                        Dim methodInfo = sb.PreCompiler.GetMethodInfo(doc.ControlsInfo(obj), method)
                        Dim ModuleName = methodInfo.Module
                        If ModuleName = "" Then
                            errors(i) = New [Error](err.Line, methodPos, $"`{match.Value}` doesn't exist.")
                            Continue For
                        End If

                        Select Case methodInfo.ParamsCount
                            Case 1
                                lines(lineNum) =
                                            prevText &
                                            $"{ModuleName}.{method}({obj})" &
                                            nextText.Substring(match.Length)
                            Case 2
                                lines(lineNum) =
                                            prevText &
                                            $"{ModuleName}.{method}({doc.Form}, {obj})" &
                                            nextText.Substring(match.Length)
                            Case Else
                                errors(i) = New [Error](err.Line, methodPos, $"`{method}` definition is not supported.")
                                Continue For
                        End Select

                        errors.RemoveAt(i)
                        ReRun = True
                    End If

                End If

            Next

            If ReRun Then Return RunProgram(
                        String.Join(Environment.NewLine, lines),
                        errors,
                        outputFileName)

            Return errors.Count = 0

        End Function

        Private Function GetBalancedBrackets(str As String) As List(Of Content)
            Do
                Dim contents As List(Of Content) = Verex.BalancedContents(str, "(", ")")
                If contents IsNot Nothing Then
                    Return contents
                End If
                Dim pos As Integer = str.LastIndexOfAny(New Char(1) {"("c, ")"c})
                If pos = -1 Then
                    Exit Do
                End If
                str = str.Substring(0, pos)
            Loop

            Return Nothing
        End Function

        Private Function CountLines(str As String) As Integer
            If str = "" Then Return 0

            Dim pos As Integer = -1
            Dim count As Integer = 0
            Do
                pos = str.IndexOf(Environment.NewLine, pos + 1)
                If pos = -1 Then Exit Do
                count += 1
            Loop
            Return count
        End Function


        Public Shared Function Compile(programText As String, outputFilePath As String) As List(Of [Error])
            Dim errors As List(Of [Error])
            Try
                Dim compiler1 As New Compiler
                compiler1.Initialize()
                Dim fileNameWithoutExtension As String = Path.GetFileNameWithoutExtension(outputFilePath)
                Dim directoryName As String = Path.GetDirectoryName(outputFilePath)
                errors = compiler1.Build(New StringReader(programText), fileNameWithoutExtension, directoryName)
            Catch ex As Exception
                errors = New List(Of [Error])
                errors.Add(New [Error](-1, 0, ex.Message))
            End Try

            Return errors
        End Function

        Private Sub formDesigner_DiagramDoubleClick(control As UIElement)
            tabCode.IsSelected = True
            If formDesigner.CodeFilePath <> "" Then
                Dim doc = OpenDocIfNot(formDesigner.CodeFilePath)
                Dim controlName = formDesigner.GetControlNameOrDefault(control)
                If doc.AddEventHandler(controlName, "OnClick") Then
                    doc.PageKey = formDesigner.PageKey
                    ' The code behind is saved before the new Handler is added.
                    ' We must make the designer dirty, to force saving this chamge in 
                    ' a nex call to SaveDesignInfo()
                    formDesigner.HasChanges = True
                End If
            End If
        End Sub

        Dim FirstTime As Boolean = True
        Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)
            If FirstTime Then
                FirstTime = False
                DiagramHelper.Designer.PagesGrid = DesignerGrid
                UpdateTitle()

                formDesigner.SavePage =
                    Function()
                        SaveDesignInfo(Nothing, False)
                        Return True
                    End Function

                AddHandler DiagramHelper.Designer.PageShown, AddressOf formDesigner_CurrentPageChanged


                'DiagramHelper.Designer.SetDefaultPropertiesSub =
                '    Sub()
                '        DiagramHelper.Designer.SetDefaultProperties()
                '    End Sub


            End If
        End Sub

        Private Sub formDesigner_CurrentPageChanged(index As Integer)
            formDesigner = DiagramHelper.Designer.CurrentPage
            ZoomBox.Designer = formDesigner
            ProjExplorer.Designer = formDesigner

            ' Remove the handler if exists not to be called twice
            RemoveHandler formDesigner.DiagramDoubleClick, AddressOf formDesigner_DiagramDoubleClick
            AddHandler formDesigner.DiagramDoubleClick, AddressOf formDesigner_DiagramDoubleClick

            UpdateTitle()

            If index > -1 AndAlso ProjExplorer.FilesList IsNot Nothing Then
                ProjExplorer.FreezListFiles = True
                ProjExplorer.FilesList.SelectedIndex = index
                ProjExplorer.FreezListFiles = False
            End If
        End Sub

        Private Sub UpdateTitle()

            If tabDesigner.IsSelected Then
                txtTitle.Text = "Form Designer - "
                txtForm.Text = $"{formDesigner.Name}{If(formDesigner.FileName = "", " *", ".xaml")}"
            Else
                txtTitle.Text = "Code Editor - "
                txtForm.Text = $"{formDesigner.Name}.sb"
            End If

        End Sub

        Dim OpeningDoc As Boolean

        Private Sub viewsControl_ActiveDocumentChanged()
            If OpeningDoc Then Return

            Dim currentView = viewsControl.SelectedItem
            If currentView Is Nothing Then Return

            For Each view As MdiView In Me.viewsControl.Items
                view.Document.OpenedInDesigner = view Is currentView
            Next

            Dim doc = currentView.Document

            If doc.PageKey = "" Then
                If doc.FilePath = "" Then
                    ' Open new page in the designer
                    doc.PageKey = DiagramHelper.Designer.OpenNewPage()
                Else
                    Dim pagePath = doc.FilePath.Substring(0, doc.FilePath.Length - 3) & ".xaml"
                    If File.Exists(pagePath) Then
                        doc.PageKey = DiagramHelper.Designer.SwitchTo(pagePath)
                    Else
                        ' Do nothing to allow opening old sb files without a form
                        ' If you want to attach a form, comment the next line,
                        ' ' and  umcomment the 2 lineslines after.

                        doc.OpenedInDesigner = False

                        ' Open new page in the designer
                        ' doc.PageKey = formDesigner.OpenNewPage()
                    End If
                End If
            Else
                DiagramHelper.Designer.SwitchTo(doc.PageKey)
            End If


        End Sub

        Private Sub ProjExplorer_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs)
            Dim item As ListBoxItem = GetParent(Of ListBoxItem)(e.OriginalSource)
            If item IsNot Nothing Then
                tabCode.IsSelected = True
            End If
        End Sub

        Private Sub CloseButton_Click(sender As Object, e As RoutedEventArgs)
            If tabCode.IsSelected Then
                If ActiveDocument IsNot Nothing Then DoCloseDoc(ActiveDocument.MdiView)
            Else
                DiagramHelper.Designer.ClosePage()
            End If

        End Sub
    End Class


End Namespace
