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
Imports RegexBuilder.Patterns
Imports sb = SmallBasic.WinForms

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

        Public ReadOnly Property DocumentTracker As DocumentTracker
            Get
                Return documentTrackerField
            End Get
        End Property

        Public ReadOnly Property ActiveDocument As TextDocument
            Get
                Dim result As TextDocument = Nothing
                Dim selectedItem As MdiView = Me.viewsControl.SelectedItem

                If selectedItem IsNot Nothing Then
                    result = selectedItem.Document
                End If

                Return result
            End Get
        End Property

        Public Sub New()
            Me.InitializeComponent()

            Try
                Dim version As Version = Assembly.GetExecutingAssembly().GetName().Version
                Me.versionText.Text = String.Format(CultureInfo.CurrentUICulture, "Microsoft Small Basic v{0}.{1}", New Object(1) {version.Major, version.Minor})
            Catch
            End Try

            mdiViews = New ObservableCollection(Of MdiView)()
            Me.viewsControl.ItemsSource = mdiViews
            AddHandler CompilerService.CurrentCompletionItemChanged, AddressOf OnCurrentCompletionItemChanged
            OnFileNew(Me, Nothing)
            helpUpdateTimer = New DispatcherTimer(TimeSpan.FromMilliseconds(200.0), DispatcherPriority.ApplicationIdle, AddressOf OnHelpUpdate, Dispatcher)
            ThreadPool.QueueUserWorkItem(New WaitCallback(AddressOf OnCheckVersion))
            Dim commandLineArgs As String() = Environment.GetCommandLineArgs()

            If commandLineArgs.Length <= 1 Then
                Return
            End If

            Dim text = commandLineArgs(1)

            If Not text.StartsWith("/") Then
                If File.Exists(text) Then
                    OpenFile(text)
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
                OpenFile(openFileDialog.FileName)
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

        Private Sub OnProgramRun(ByVal sender As Object, ByVal e As RoutedEventArgs)
            RunProgram(ActiveDocument)
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

        Private Sub OpenFile(ByVal filePath As String)
            Dim document As New TextDocument(filePath)
            DocumentTracker.TrackDocument(document)
            Dim mdiView As MdiView = New MdiView()
            mdiView.Document = document
            Dim item = mdiView
            mdiViews.Add(item)
        End Sub

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


        Private Function RunProgram(ByVal document As TextDocument) As Boolean
            Dim outputFileName As String

            Try
                outputFileName = GetOutputFileName(document)
            Catch ex As Exception
                Utility.MessageBox.Show(ResourceHelper.GetString("FailedToCreateOutputFile"), ResourceHelper.GetString("Title"), String.Format(CultureInfo.CurrentUICulture, ResourceHelper.GetString("FailedToCreateOutputFileReason"), New Object(0) {ex.Message}), NotificationButtons.Close, NotificationIcon.Error)
                Return True
            End Try

            document.Errors.Clear()

            If CompilerService.Compile(document.Text, outputFileName, document.Errors) Then
                Thread.Sleep(500)
                currentProgramProcess = Process.Start(outputFileName)
                currentProgramProcess.EnableRaisingEvents = True
                AddHandler currentProgramProcess.Exited,
                    Sub()
                        Dispatcher.BeginInvoke(DispatcherPriority.Normal, CType(
                             Function()
                                 Me.programRunningOverlay.Visibility = Visibility.Hidden
                                 currentProgramProcess = Nothing

                                 If document.IsNew Then
                                     Try
                                         File.Delete(outputFileName)
                                     Catch
                                     End Try
                                 End If

                                 ActiveDocument.EditorControl.Focus()
                                 Return Nothing
                             End Function,
                             DispatcherOperationCallback), Nothing)
                    End Sub

                Me.processRunningMessage.Text = String.Format(ResourceHelper.GetString("ProgramRunning"), document.Title)
                Me.programRunningOverlay.Visibility = Visibility.Visible
                Me.endProgramButton.Focus()

            ElseIf document.ParseFormHints() Then
                Return PreCompile(document)
            End If

            Return True
        End Function

        Private ReadOnly WordRgex As New Verex(Symbols.AnyWord)
        Private ReadOnly OpenBracketRegex As New Verex(NoneOrMany(Symbols.WhiteSpace) + "(")
        Private ReadOnly MethodRegex As New Verex(Symbols.AnyWord + NoneOrMany(Symbols.WhiteSpace) + "(")

        Private Function PreCompile(document As TextDocument) As Boolean
            Dim ReRun = False
            Dim txt = document.Text
            Dim lines = txt.Split({Environment.NewLine}, StringSplitOptions.None)

            For i = document.Errors.Count - 1 To 0 Step -1
                Dim err = document.Errors(i)

                Dim pos1 = err.IndexOf(",")
                If pos1 = -1 Then Continue For

                Dim lineNum = CInt(err.Substring(0, pos1)) - 1
                pos1 += 1
                Dim pos2 = err.IndexOf(":", pos1)
                Dim CharNum = CInt(err.Substring(pos1, pos2 - pos1)) - 1
                Dim errMsg = err.Substring(pos2 + 2)
                If errMsg.StartsWith("Cannot find object") Then
                    pos1 = errMsg.LastIndexOf("'", errMsg.Length - 3) + 1
                    Dim obj = errMsg.Substring(pos1, errMsg.Length - pos1 - 2)
                    If Not document.ControlsInfo.ContainsKey(obj) Then Continue For
                    Dim line = lines(lineNum)

                    If line.Substring(CharNum, obj.Length + 1) = obj + "." Then
                        Dim prevText = If(CharNum = 0, "", line.Substring(0, CharNum))
                        Dim nextText = line.Substring(CharNum + obj.Length + 1)

                        Dim match = MethodRegex.Match(nextText)
                        If match.Success AndAlso match.Index = 0 Then ' Method Call
                            Dim method = nextText.Substring(0, WordRgex.Match(nextText).Length)
                            Dim contents = GetBalancedBrackets(nextText)
                            If contents Is Nothing OrElse contents.Count = 0 Then Continue For
                            Dim params = contents(0).Value.Trim(" "c, Convert.ToChar(8))
                            Dim RestText = If(contents(0).Index + contents(0).Length > nextText.Length - 2, "",
                                        nextText.Substring(contents(0).Index + contents(0).Length + 1))

                            Dim ModuleName = sb.PreCompiler.GetModule(document.ControlsInfo(obj), method)
                            If ModuleName = "" Then Continue For

                            If params = "" Then
                                lines(lineNum) = prevText &
                                    $"{ModuleName}.{method}({document.Form}, {obj})" &
                                    RestText
                            Else
                                lines(lineNum) = prevText &
                                    $"{ModuleName}.{method}({document.Form}, {obj}, {params})" &
                                    RestText
                            End If

                            document.Errors.RemoveAt(i)
                            ReRun = True
                        ElseIf prevText.Trim(" "c, Convert.ToChar(8)) = "" Then 'Property Set
                            pos1 = nextText.IndexOf("="c)
                            If pos1 = -1 Then Continue For
                            Dim result = WordRgex.Match(nextText)
                            If Not result.Success OrElse result.Index > 0 Then Continue For
                            Dim L = pos1 - result.Length

                            If L = 0 OrElse nextText.Substring(result.Length, L).Trim(" "c, Convert.ToChar(8)) = "" Then
                                Dim method = $"Set{result.Value}"
                                Dim ModuleName = sb.PreCompiler.GetModule(document.ControlsInfo(obj), method)
                                If ModuleName = "" Then Continue For

                                lines(lineNum) = prevText &
                                    $"{ModuleName}.{method}({document.Form}, {obj}, {nextText.Substring(pos1 + 1).Trim})"
                                document.Errors.RemoveAt(i)
                                ReRun = True
                            End If

                        Else 'Property Get          
                                match = WordRgex.Match(nextText)
                            If Not match.Success OrElse match.Index > 0 Then Continue For

                            Dim method = $"Get{match.Value}"
                            Dim ModuleName = sb.PreCompiler.GetModule(document.ControlsInfo(obj), method)
                            If ModuleName = "" Then Continue For

                            lines(lineNum) =
                                    prevText &
                                    $"{ModuleName}.{method}({document.Form}, {obj})" &
                                    nextText.Substring(match.Length)
                            document.Errors.RemoveAt(i)
                            ReRun = True
                        End If

                    End If

                End If
            Next

            If ReRun Then Return RunAgain(document, lines)

            Return document.Errors.Count = 0

        End Function

        Private Function GetBalancedBrackets(str As String) As List(Of Content)
            Do
                Dim contents = Verex.BalancedContents(str, "(", ")")
                If contents IsNot Nothing Then Return contents
                Dim pos = str.LastIndexOfAny({"("c, ")"c})
                If pos = -1 Then Return Nothing
                str = str.Substring(0, pos)
            Loop
            Return Nothing
        End Function

        Function RunAgain(document As TextDocument, lines() As String) As Boolean
            Dim n = New Random().Next(1, 1000000)
            Dim filename = Path.Combine(Path.GetTempPath(), $"file{n}.sb")
            My.Computer.FileSystem.WriteAllText(filename, String.Join(Environment.NewLine, lines), False)
            Dim doc As New TextDocument(filename)

            document.Errors.Clear()
            If RunProgram(doc) Then Return True

            For Each err In doc.Errors
                document.Errors.Add(err)
            Next


            Return False

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
                        Dim item = mdiView
                        mdiViews.Add(item)
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

        Private Sub OnCheckVersion(ByVal state As Object)
            Thread.Sleep(20000)

            Try
                Dim service As Service = New Service()
                Dim currentVersion As String = service.GetCurrentVersion()
                Dim version As Version = New Version(currentVersion)
                Dim version2 As Version = Assembly.GetExecutingAssembly().GetName().Version

                If version.CompareTo(version2) > 0 Then
                    Dispatcher.BeginInvoke(CType(Sub() Me.updateAvailable.Visibility = Visibility.Visible, Action))
                End If

            Catch
            End Try
        End Sub

        Private Sub OnClickNewVersionAvailable(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Process.Start("http://smallbasic.com/download.aspx")
        End Sub

    End Class
End Namespace
