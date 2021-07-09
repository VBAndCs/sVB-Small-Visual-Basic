Imports Microsoft.Nautilus.Text
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

            'Try
            '    Dim version As Version = Assembly.GetExecutingAssembly().GetName().Version
            '    Me.versionText.Text = String.Format(CultureInfo.CurrentUICulture, "Microsoft Small Basic v{0}.{1}", New Object(1) {version.Major, version.Minor})
            'Catch
            'End Try

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
                    OpenDocumentIfNot(text)
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
                OpenDocumentIfNot(openFileDialog.FileName)
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

        Private Sub OnCloseItem(ByVal sender As Object, ByVal e As RequestCloseEventArgs)
            Dim mdiView As MdiView = TryCast(e.Item, MdiView)

            If mdiView IsNot Nothing AndAlso CloseDocument(mdiView.Document) Then
                mdiViews.Remove(mdiView)
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

            Dim textBox As New TextBox()
            textBox.Text = text
            textBox.FontSize = 20.0
            textBox.FontFamily = New FontFamily("Consolas")
            textBox.Foreground = Brushes.DimGray
            textBox.Margin = New Thickness(0.0, 4.0, 4.0, 4.0)
            textBox.MinWidth = 300.0
            messageBox.OptionalContent = textBox
            Dim textBox3 = textBox
            messageBox.NotificationButtons = Nf.Close Or Nf.OK
            messageBox.okButton.Content = ResourceHelper.GetString("FindCommand")
            messageBox.NotificationIcon = NotificationIcon.Custom
            messageBox.iconImageControl.Source = New BitmapImage(New Uri("pack://application:,,/SB;component/Resources/Search.png"))
            textBox3.SelectAll()
            textBox3.Focus()
            Dim notificationButtons As NotificationButtons = messageBox.Display()

            If notificationButtons = NotificationButtons.OK Then
                lastSearchedText = textBox3.Text

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
                    Exit For
                End If

                mdiViews.Remove(item)
            Next

            MyBase.OnClosing(e)
        End Sub

        Private Function OpenFile(ByVal filePath As String) As TextDocument
            Dim document As New TextDocument(filePath)
            DocumentTracker.TrackDocument(document)
            Dim mdiView As New MdiView()
            mdiView.Document = document
            mdiViews.Add(mdiView)
            Return document
        End Function

        Private Function SaveDocument(ByVal document As TextDocument) As Boolean
            If document.IsNew Then
                Return SaveDocumentAs(document)
            End If

            Try
                document.Save()
                Return True
            Catch ex As Exception
                Dim notificationButtons = Utility.MessageBox.Show(ResourceHelper.GetString("SaveFailed"), "Small Basic", String.Format(CultureInfo.CurrentCulture, ResourceHelper.GetString("SaveFailedReason"), New Object(0) {ex.Message}), Nf.No Or Nf.Yes, NotificationIcon.Information)

                If notificationButtons = NotificationButtons.Yes Then
                    Return SaveDocumentAs(document)
                End If

                Return False
            End Try
        End Function

        Private Function SaveDocumentAs(ByVal document As TextDocument) As Boolean
            Dim saveFileDialog As SaveFileDialog = New SaveFileDialog()
            saveFileDialog.Filter = ResourceHelper.GetString("SmallBasicFileFilter") & "|*.sb"

            If saveFileDialog.ShowDialog() = True Then
                Try
                    document.SaveAs(saveFileDialog.FileName)
                    Return True
                Catch ex As Exception
                    Dim notificationButtons = Utility.MessageBox.Show(ResourceHelper.GetString("SaveFailed"), ResourceHelper.GetString("Title"), String.Format(CultureInfo.CurrentUICulture, ResourceHelper.GetString("SaveFailedReason"), New Object(0) {ex.Message}), Nf.No Or Nf.Yes, NotificationIcon.Information)

                    If notificationButtons = NotificationButtons.Yes Then
                        Return SaveDocumentAs(document)
                    End If

                    Return False
                End Try
            End If

            Return False
        End Function

        Private Function CloseDocument(ByVal document As TextDocument) As Boolean
            If document.IsDirty Then
                Select Case Utility.MessageBox.Show(ResourceHelper.GetString("SaveDocumentBeforeClosing"), ResourceHelper.GetString("Title"), document.Title & ResourceHelper.GetString("DocumentModified"), NotificationButtons.Cancel Or NotificationButtons.No Or NotificationButtons.Yes, NotificationIcon.Information)
                    Case NotificationButtons.Yes
                        Return SaveDocument(document)
                    Case NotificationButtons.No
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
                If doc.Form <> "" Then SaveDesignInfo(doc)

                Dim code = ""
                Dim outputFileName = GetOutputFileName(doc)
                Dim genCodefile = doc.FilePath
                Dim offset = 0
                genCodefile = genCodefile?.Substring(0, genCodefile.Length - 2) + "gsb"

                If IO.File.Exists(genCodefile) Then
                    Dim gen = File.ReadAllText(genCodefile)
                    offset = CountLines(gen) + 1
                    code = gen & Environment.NewLine & doc.Text
                    If doc.Form = "" Then doc.ParseFormHints(code)
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

                                 doc.EditorControl.Focus()
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
                Dim messageBox As Utility.MessageBox = New Utility.MessageBox()
                messageBox.Description = ResourceHelper.GetString("ImportFromWeb")
                messageBox.Title = ResourceHelper.GetString("Title")
                Dim stackPanel As StackPanel = New StackPanel()
                stackPanel.Orientation = Orientation.Vertical
                Dim stackPanel2 = stackPanel
                Dim textBlock As TextBlock = New TextBlock()
                textBlock.Text = ResourceHelper.GetString("ImportLocationOfProgramOnWeb")
                textBlock.Margin = New Thickness(0.0, 0.0, 4.0, 4.0)
                Dim element = textBlock
                Dim textBox As TextBox = New TextBox()
                textBox.FontSize = 32.0
                textBox.FontWeight = FontWeights.Bold
                textBox.FontFamily = New FontFamily("Courier New")
                textBox.Foreground = Brushes.DimGray
                textBox.Margin = New Thickness(0.0, 4.0, 4.0, 4.0)
                textBox.MinWidth = 300.0
                Dim textBox2 = textBox
                stackPanel2.Children.Add(element)
                stackPanel2.Children.Add(textBox2)
                messageBox.OptionalContent = stackPanel2
                messageBox.NotificationButtons = NotificationButtons.Cancel Or NotificationButtons.OK
                messageBox.NotificationIcon = NotificationIcon.Information
                textBox2.Focus()

                If messageBox.Display() = NotificationButtons.OK Then
                    Dim service As Service = New Service()
                    Dim text As String = textBox2.Text.Trim()
                    Dim text2 = service.LoadProgram(text)

                    If Equals(text2, "error") Then
                        Utility.MessageBox.Show(ResourceHelper.GetString("FailedToImportFromWeb"), ResourceHelper.GetString("Title"), ResourceHelper.GetString("ImportFromWebFailedReason"), NotificationButtons.Close, NotificationIcon.Error)
                    Else
                        text2 = text2.Replace(VisualBasic.Constants.vbLf, VisualBasic.Constants.vbCrLf)
                        Dim newDocument As New TextDocument(Nothing)
                        newDocument.ContentType = "text.smallbasic"
                        newDocument.BaseId = text
                        newDocument.TextBuffer.Insert(0, text2)
                        Dim service2 As Service = New Service()
                        AddHandler service2.GetProgramDetailsCompleted, Sub(ByVal o, ByVal e)
                                                                            Dim result = e.Result
                                                                            result.Category = ResourceHelper.GetString("Category" & result.Category)
                                                                            newDocument.ProgramDetails = result
                                                                        End Sub

                        service2.GetProgramDetailsAsync(text)
                        DocumentTracker.TrackDocument(newDocument)
                        Dim mdiView As MdiView = New MdiView()
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

        Private Function GetOutputFileName(ByVal document As TextDocument) As String
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

        Function OpenDocumentIfNot(FilePath As String) As TextDocument
            For Each view As MdiView In Me.viewsControl.Items
                If view.Document.FilePath = Path.GetFullPath(FilePath) Then
                    viewsControl.ChangeSelection(view)
                    Return view.Document
                End If
            Next
            Return OpenFile(FilePath)
        End Function

        Private Sub tabCode_Selected(sender As Object, e As RoutedEventArgs)
            SaveDesignInfo()
        End Sub

        Private Sub SaveDesignInfo(Optional doc As TextDocument = Nothing)
            Dim hint As New Text.StringBuilder
            Dim declaration As New Text.StringBuilder
            Dim formName As String
            Dim xamlPath As String
            Dim formPath As String

            If formDesigner.FileName = "" Then
                Dim tmpPath = "UnSaved"
                If Not IO.Directory.Exists(tmpPath) Then IO.Directory.CreateDirectory(tmpPath)

                Dim fileName = ""
                Dim projectName = ""
                Dim projectPath = ""
                Dim n = 1
                Do
                    projectName = "Project" & n
                    projectPath = Path.Combine(tmpPath, projectName)
                    If Not IO.Directory.Exists(projectPath) Then
                        IO.Directory.CreateDirectory(projectPath)
                        Exit Do
                    End If
                    n += 1
                Loop

                n = 1
                xamlPath = ""
                Do
                    formName = "Form" & n
                    xamlPath = Path.Combine(projectPath, formName)
                    If Not IO.Directory.Exists(xamlPath) Then
                        IO.Directory.CreateDirectory(xamlPath)
                        Exit Do
                    End If
                    n += 1
                Loop

                formPath = Path.Combine(xamlPath, formName)
                formDesigner.FileName = formPath & ".xaml"
                formDesigner.DoSave()
                IO.File.Create(formPath & ".sb").Close()

            Else
                formName = Path.GetFileNameWithoutExtension(formDesigner.FileName)
                xamlPath = Path.GetDirectoryName(formDesigner.FileName)
                formPath = Path.Combine(xamlPath, formName)
                If formDesigner.HasChanges Then formDesigner.DoSave()
            End If

            hint.AppendLine($"'#{formName}{{")
            Dim ControlsInfo As New Dictionary(Of String, String)
            Dim ControlNames As New List(Of String)
            ControlNames.Add("(Global)")

            ControlsInfo(formName.ToLower()) = "Form"
            ControlsInfo("me") = "Form"
            ControlNames.Add(formName)
            declaration.AppendLine($"Me = ""{formName}""")

            For Each c As FrameworkElement In formDesigner.Items
                Dim name = c.Name
                If name <> "" Then
                    Dim typeName = c.GetType().Name
                    ControlsInfo(name.ToLower()) = typeName
                    ControlNames.Add(name)
                    hint.AppendLine($"'    {name}: {typeName}")
                    declaration.AppendLine($"{name} = ""{name}""")
                End If
            Next

            hint.AppendLine("'}")
            hint.AppendLine()
            hint.Append(declaration)
            hint.AppendLine("True = 1")
            hint.AppendLine("False = 0")
            hint.AppendLine($"Forms.AppPath = ""{xamlPath}""")
            hint.AppendLine($"{formName} = Forms.LoadForm(""{formName}"", ""{formName}.xaml"")")
            hint.AppendLine($"Control.SetWidth({formName}, {formName}, {formDesigner.PageWidth})")
            hint.AppendLine($"Control.SetHeight({formName}, {formName}, {formDesigner.PageHeight})")
            hint.AppendLine($"Form.Show({formName})")

            If doc Is Nothing Then doc = OpenDocumentIfNot(formPath & ".sb")

            doc.RemoveBrokenHandlers()

            If doc.EventHandlers.Count > 0 Then
                hint.AppendLine()
                hint.AppendLine("'#Events{")
                Dim sbHandlers As New Text.StringBuilder

                Dim controlEvents = From eventHandler In doc.EventHandlers
                                    Group By eventHandler.Value.ControlName
                           Into EventInfo = Group

                For Each ev In controlEvents
                    hint.Append($"'    {ev.ControlName}:")
                    sbHandlers.AppendLine($"' {ev.ControlName} Events:")
                    sbHandlers.AppendLine($"Control.HandleEvents({formName}, {ev.ControlName})")
                    For Each info In ev.EventInfo
                        hint.Append($" {info.Value.EventName}")
                        Dim module1 = ControlsInfo(ev.ControlName.ToLower)
                        Dim moduleName = sb.PreCompiler.GetEventModule(module1, info.Value.EventName)
                        sbHandlers.AppendLine($"{moduleName}.{info.Value.EventName} = {info.Key}")
                    Next
                    hint.AppendLine()
                    sbHandlers.AppendLine()
                Next
                hint.AppendLine("'}")
                hint.AppendLine()
                hint.AppendLine(sbHandlers.ToString())
            End If

            IO.File.WriteAllText(formPath & ".gsb", hint.ToString())

            doc.Form = formName.ToLower()
            doc.ControlsInfo = ControlsInfo

            ' Note that ControlNames Property is bound to a cono box, so keep the existing collection
            ControlNames.Sort()
            doc.ControlNames.Clear()
            For Each controlName In ControlNames
                doc.ControlNames.Add(controlName)
            Next
        End Sub

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
                If Not errMsg.StartsWith("Cannot find object") Then Continue For

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

    End Class


End Namespace
