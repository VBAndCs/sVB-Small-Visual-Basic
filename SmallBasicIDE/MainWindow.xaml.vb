Imports Microsoft.Nautilus.Text
Imports Microsoft.SmallVisualBasic.com.smallbasic
Imports Microsoft.SmallVisualBasic.Documents
Imports Microsoft.SmallVisualBasic.LanguageService
Imports Microsoft.SmallVisualBasic.Shell
Imports Microsoft.SmallVisualBasic.Utility
Imports Microsoft.Win32
Imports Microsoft.Windows.Controls
Imports System
Imports System.Collections.Generic
Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.Diagnostics
Imports System.Globalization
Imports System.IO
Imports Nf = Microsoft.SmallVisualBasic.Utility.NotificationButtons
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
Imports sb = Microsoft.SmallVisualBasic.WinForms
Imports Microsoft.VisualBasic
Imports Microsoft.SmallBasic

Namespace Microsoft.SmallVisualBasic
    <Export("MainWindow")>
    Public Class MainWindow
        Private currentProgramProcess As Process
        Private ReadOnly mdiViews As ObservableCollection(Of MdiView)
        Private lastSearchedText As String = ""

        Public Shared NewCommand As New RoutedUICommand(
            ResourceHelper.GetString("NewCommand"),
            ResourceHelper.GetString("NewCommand"),
            GetType(MainWindow)
        )

        Public Shared OpenCommand As New RoutedUICommand(
            ResourceHelper.GetString("OpenCommand"),
            ResourceHelper.GetString("OpenCommand"),
            GetType(MainWindow)
        )

        Public Shared SaveCommand As New RoutedUICommand(
            ResourceHelper.GetString("SaveCommand"),
            ResourceHelper.GetString("SaveCommand"),
            GetType(MainWindow)
        )

        Public Shared SaveAsCommand As New RoutedUICommand(
            ResourceHelper.GetString("SaveAsCommand"),
            ResourceHelper.GetString("SaveAsCommand"),
            GetType(MainWindow)
        )

        Public Shared CutCommand As New RoutedUICommand(
            ResourceHelper.GetString("CutCommand"),
            ResourceHelper.GetString("CutCommand"),
            GetType(MainWindow)
        )

        Public Shared CopyCommand As New RoutedUICommand(
            ResourceHelper.GetString("CopyCommand"),
            ResourceHelper.GetString("CopyCommand"),
            GetType(MainWindow)
         )

        Public Shared PasteCommand As New RoutedUICommand(
            ResourceHelper.GetString("PasteCommand"),
            ResourceHelper.GetString("PasteCommand"),
            GetType(MainWindow)
        )

        Public Shared FindCommand As New RoutedUICommand(
            ResourceHelper.GetString("FindCommand"),
            ResourceHelper.GetString("FindCommand"),
            GetType(MainWindow)
        )

        Public Shared FindNextCommand As New RoutedUICommand(
            ResourceHelper.GetString("FindNextCommand"),
            ResourceHelper.GetString("FindNextCommand"),
            GetType(MainWindow)
        )

        Public Shared SelectNextMatchCommand As New RoutedUICommand(
            "Find Next matching pair",
            "Find Next matching pair",
            GetType(MainWindow)
        )

        Public Shared SelectPrevMatchCommand As New RoutedUICommand(
            "Find Previous matching pair",
            "Find Previous matching pair",
            GetType(MainWindow)
        )

        Public Shared UndoCommand As New RoutedUICommand(
            ResourceHelper.GetString("UndoCommand"),
            ResourceHelper.GetString("UndoCommand"),
            GetType(MainWindow)
        )

        Public Shared RedoCommand As New RoutedUICommand(
            ResourceHelper.GetString("RedoCommand"),
            ResourceHelper.GetString("RedoCommand"),
            GetType(MainWindow)
        )

        Public Shared FormatCommand As New RoutedUICommand(
            ResourceHelper.GetString("FormatProgramCommand"),
            ResourceHelper.GetString("FormatProgramCommand"),
            GetType(MainWindow)
        )

        Public Shared RunCommand As New RoutedUICommand(
            ResourceHelper.GetString("RunProgramCommand"),
            ResourceHelper.GetString("RunProgramCommand"),
            GetType(MainWindow)
        )

        Public Shared EndProgramCommand As New RoutedUICommand(
            ResourceHelper.GetString("EndProgramCommand"),
            ResourceHelper.GetString("EndProgramCommand"),
            GetType(MainWindow)
        )

        Public Shared StepOverCommand As New RoutedUICommand(
            ResourceHelper.GetString("StepOverCommand"),
            ResourceHelper.GetString("StepOverCommand"),
            GetType(MainWindow)
        )

        Public Shared BreakpointCommand As New RoutedUICommand(
            ResourceHelper.GetString("BreakpointCommand"),
            ResourceHelper.GetString("BreakpointCommand"),
            GetType(MainWindow)
         )

        Public Shared DebugCommand As New RoutedUICommand(
            ResourceHelper.GetString("DebugCommand"),
            ResourceHelper.GetString("DebugCommand"),
            GetType(MainWindow)
         )

        Public Shared WebSaveCommand As New RoutedUICommand(
            ResourceHelper.GetString("PublishProgramCommand"),
            ResourceHelper.GetString("PublishProgramCommand"),
            GetType(MainWindow)
        )

        Public Shared WebLoadCommand As New RoutedUICommand(
            ResourceHelper.GetString("ImportProgramCommand"),
            ResourceHelper.GetString("ImportProgramCommand"),
            GetType(MainWindow)
        )

        Public Shared ExportToVisualBasicCommand As New RoutedUICommand(
            ResourceHelper.GetString("ExportToVisualBasicCommand"),
            ResourceHelper.GetString("ExportToVisualBasicCommand"),
            GetType(MainWindow)
        )

        ''' <summary>
        ''' DocumentTracker
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property DocumentTracker As New DocumentTracker()

        Public ReadOnly Property ActiveDocument As TextDocument
            Get
                Dim selectedItem = viewsControl.SelectedItem
                If selectedItem IsNot Nothing Then
                    Return selectedItem.Document
                Else
                    Return Nothing
                End If
            End Get
        End Property

        Public Sub New()
            Me.InitializeComponent()
            mdiViews = New ObservableCollection(Of MdiView)()
            Me.viewsControl.ItemsSource = mdiViews
            AddHandler CompilerService.HelpUpdated,
                    Sub(wrapper) HelpPanel.CompletionItemWrapper = wrapper
        End Sub

        Private Sub OnFileNew(sender As Object, e As RoutedEventArgs)
            Dim doc As New TextDocument(Nothing) With {
                .ContentType = "text.smallbasic",
                .IsNew = True
            }
            _DocumentTracker.TrackDocument(doc)
            Dim mdiView As New MdiView() With {
                  .Document = doc,
                  .Width = viewsControl.ActualWidth - 100,
                  .Height = viewsControl.ActualHeight - 50
            }
            Canvas.SetTop(mdiView, 10)
            Canvas.SetLeft(mdiView, 10)
            mdiViews.Add(mdiView)
            doc.Focus()
        End Sub

        Private Sub OnFileOpen(sender As Object, e As RoutedEventArgs)
            Dim openFileDialog As New OpenFileDialog()
            openFileDialog.Filter = ResourceHelper.GetString("SmallBasicFileFilter") & "|*.sb;*.smallbasic"

            If openFileDialog.ShowDialog() = True Then
                OpenDocIfNot(openFileDialog.FileName).Focus()
            End If

        End Sub

        Private Sub CanFileSave(sender As Object, e As CanExecuteRoutedEventArgs)
            e.CanExecute = ActiveDocument IsNot Nothing
        End Sub

        Private Sub OnFileSave(sender As Object, e As RoutedEventArgs)
            SaveDocument(ActiveDocument)
        End Sub

        Private Sub OnFileSaveAs(sender As Object, e As RoutedEventArgs)
            SaveDocumentAs(ActiveDocument)
        End Sub

        Private Sub CanEditCut(sender As Object, e As CanExecuteRoutedEventArgs)
            e.CanExecute = ActiveDocument IsNot Nothing AndAlso Not ActiveDocument.EditorControl.TextView.Selection.IsEmpty
        End Sub

        Private Sub OnEditCut(sender As Object, e As RoutedEventArgs)
            Dim doc = ActiveDocument
            doc.EditorControl.EditorOperations.CutSelection(doc.UndoHistory)
        End Sub

        Private Sub CanEditCopy(sender As Object, e As CanExecuteRoutedEventArgs)
            e.CanExecute = ActiveDocument IsNot Nothing AndAlso Not ActiveDocument.EditorControl.TextView.Selection.IsEmpty
        End Sub


        Private Sub OnEditCopy(sender As Object, e As RoutedEventArgs)
            ActiveDocument.EditorControl.EditorOperations.CopySelection()
        End Sub

        Private Sub CanEditPaste(sender As Object, e As CanExecuteRoutedEventArgs)
            e.CanExecute = ActiveDocument IsNot Nothing AndAlso ActiveDocument.EditorControl.EditorOperations.CanPaste
        End Sub

        Private Sub OnEditPaste(sender As Object, e As RoutedEventArgs)
            Dim doc = ActiveDocument
            doc.EditorControl.EditorOperations.Paste(doc.UndoHistory)
        End Sub

        Private Sub CanEditUndo(sender As Object, e As CanExecuteRoutedEventArgs)
            e.CanExecute = ActiveDocument IsNot Nothing AndAlso ActiveDocument.UndoHistory.CanUndo
        End Sub

        Private Sub OnEditUndo(sender As Object, e As RoutedEventArgs)
            ActiveDocument.UndoHistory.Undo(1)
        End Sub

        Private Sub CanEditRedo(sender As Object, e As CanExecuteRoutedEventArgs)
            e.CanExecute = ActiveDocument IsNot Nothing AndAlso ActiveDocument.UndoHistory.CanRedo
        End Sub

        Private Sub OnEditRedo(sender As Object, e As RoutedEventArgs)
            ActiveDocument.UndoHistory.Redo(1)
        End Sub

        Private Sub OnCloseItem(sender As Object, e As RequestCloseEventArgs)
            Dim mdiView = TryCast(e.Item, MdiView)
            CloseView(mdiView)
        End Sub

        Public Sub CloseView(mdiView As MdiView)
            If mdiView IsNot Nothing AndAlso CloseDocument(mdiView.Document) Then
                mdiView.Document.Close()
                mdiViews.Remove(mdiView)
                If mdiViews.Count > 0 Then
                    If viewsControl.SelectedItem Is Nothing Then
                        mdiView = mdiViews(mdiViews.Count - 1)
                        viewsControl.ChangeSelection(mdiView)
                        mdiView.Document.Focus()
                    End If
                End If
            End If
        End Sub

        Private Sub CanExportToVisualBasic(sender As Object, e As CanExecuteRoutedEventArgs)
            If ActiveDocument IsNot Nothing AndAlso ActiveDocument.Text.Trim().Length > 0 Then
                e.CanExecute = True
            Else
                e.CanExecute = False
            End If
        End Sub

        Private Sub OnExportToVisualBasic(sender As Object, e As RoutedEventArgs)
            Dim exportToVBDialog As New ExportToVBDialog(ActiveDocument)
            exportToVBDialog.Owner = Me
            exportToVBDialog.ShowDialog()
        End Sub

        Private Sub OnFind(sender As Object, e As RoutedEventArgs)
            If ActiveDocument Is Nothing Then
                Return
            End If

            Dim messageBox As New Utility.MessageBox With {
                .Description = ResourceHelper.GetString("TextToSearch"),
                .Title = ResourceHelper.GetString("FindCommand")
            }
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
            messageBox.iconImageControl.Source = New BitmapImage(New Uri("pack://application:,,/sVB;component/Resources/Search.png"))
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

        Private Sub OnFindNext(sender As Object, e As RoutedEventArgs)
            If ActiveDocument IsNot Nothing Then
                Dim editor = ActiveDocument.EditorControl
                If Not editor.ContainsHighlights AndAlso lastSearchedText = "" Then
                    OnFind(sender, e)
                Else
                    Dim srch = If(lastSearchedText = "",
                         editor.TextView.Selection?.ActiveSnapshotSpan.GetText(),
                         lastSearchedText
                    )

                    If Not editor.HighlightNextMatch(srch, ignoreCase:=True) Then Console.Beep()
                End If
            End If
        End Sub

        Friend Sub SelectControl(controlName As String)
            controlName = controlName.ToLower()
            If controlName = "me" OrElse controlName = formDesigner.Name.ToLower() Then
                tabDesigner.IsSelected = True
                formDesigner.SelectedIndex = -1
                formDesigner.Focus()
            Else
                For i = 0 To formDesigner.Items.Count - 1
                    Dim name = formDesigner.GetControlName(i).ToLower()
                    If controlName = name Then
                        tabDesigner.IsSelected = True
                        formDesigner.SelectedIndex = i
                        Return
                    End If
                Next
                Beep()
            End If

        End Sub

        Private Sub OnSelectNextMatch(sender As Object, e As RoutedEventArgs)
            SelectAnotherMatch(True)
        End Sub

        Private Sub OnSelectPrevMatch(sender As Object, e As RoutedEventArgs)
            SelectAnotherMatch(False)
        End Sub

        Private Sub SelectAnotherMatch(moveDown As Boolean)
            Dim doc = ActiveDocument
            If doc IsNot Nothing Then
                Dim editor = doc.EditorControl
                If Not editor.ContainsWordHighlights Then
                    doc.HighlightEnclosingBlockKeywords()
                End If
                editor.SelectAnotherHighlightedWord(moveDown)
            End If
        End Sub

        Private Sub OnFormatProgram(sender As Object, e As RoutedEventArgs)
            Dim doc = ActiveDocument
            Format(doc)
        End Sub

        Sub Format(doc As TextDocument)
            If doc IsNot Nothing Then
                Using undoTransaction = doc.UndoHistory.CreateTransaction("Format Document")
                    FormatDocument(doc.TextBuffer)
                    undoTransaction.Complete()
                End Using
            End If
        End Sub

        Private Sub CanRunProgram(sender As Object, e As CanExecuteRoutedEventArgs)
            e.CanExecute = ActiveDocument IsNot Nothing
        End Sub

        Private Sub OnProgramRun(sender As Object, e As RoutedEventArgs)
            RunProgram()
        End Sub

        Private Sub OnProgramEnd(sender As Object, e As RoutedEventArgs)
            If currentProgramProcess IsNot Nothing AndAlso Not currentProgramProcess.HasExited Then
                currentProgramProcess.Kill()
                currentProgramProcess = Nothing
                Me.programRunningOverlay.Visibility = Visibility.Hidden
            End If
        End Sub

        Private Sub OnStepOver(sender As Object, e As RoutedEventArgs)
            ActiveDocument.Errors.Clear()
            Compile(ActiveDocument.Text, ActiveDocument.Errors)

            If ActiveDocument.Errors.Count = 0 Then
                Dim debugger = ProgramDebugger.GetDebugger(ActiveDocument)
                debugger.StepOver()
            End If
        End Sub

        Private Sub OnToggleBreakpoint(sender As Object, e As RoutedEventArgs)
            Dim debugger = ProgramDebugger.GetDebugger(ActiveDocument)
            Dim stMarkerProvider = StatementMarkerProvider.GetStatementMarkerProvider(ActiveDocument.EditorControl.TextView)
            Dim line = ActiveDocument.TextBuffer.CurrentSnapshot.GetLineFromPosition(ActiveDocument.EditorControl.TextView.Caret.Position.CharacterIndex)

            If debugger.Breakpoints.Contains(line.LineNumber) Then
                debugger.Breakpoints.Remove(line.LineNumber)
                Dim array As StatementMarker() = stMarkerProvider.Markers.ToArray()

                For Each statementMarker In array
                    Dim containingLine As ITextSnapshotLine = statementMarker.Span.GetStartPoint(ActiveDocument.TextBuffer.CurrentSnapshot).GetContainingLine()

                    If containingLine.LineNumber = line.LineNumber Then
                        stMarkerProvider.RemoveMarker(statementMarker)
                    End If
                Next
            Else
                debugger.Breakpoints.Add(line.LineNumber)
                stMarkerProvider.AddStatementMarker(New StatementMarker(New TextSpan(ActiveDocument.TextBuffer.CurrentSnapshot, line.Start, line.Length, SpanTrackingMode.EdgeInclusive), Colors.Red))
            End If
        End Sub

        Private Sub CanWebSave(sender As Object, e As CanExecuteRoutedEventArgs)
            e.CanExecute = ActiveDocument IsNot Nothing
        End Sub

        Private Sub OnWebSave(sender As Object, e As RoutedEventArgs)
            PublishDocument(ActiveDocument)
        End Sub

        Private Sub OnWebLoad(sender As Object, e As RoutedEventArgs)
            ImportDocument()
        End Sub


        Function CloseDocument(docIndex As Integer) As Boolean
            Dim item = mdiViews(docIndex)
            If Not CloseDocument(item.Document) Then Return False
            mdiViews.RemoveAt(docIndex)
            Return True
        End Function

        Protected Overrides Sub OnClosing(e As CancelEventArgs)
            Dim doc = ActiveDocument
            If doc IsNot Nothing Then
                If Not CloseDocument(doc) Then
                    e.Cancel = True
                    Return
                End If
                mdiViews.Remove(doc.MdiView)
            End If

            For i = mdiViews.Count - 1 To 0 Step -1
                If Not CloseDocument(i) Then
                    e.Cancel = True
                    Return
                End If
            Next

            tabDesigner.IsSelected = True
            For i = 0 To DiagramHelper.Designer.Pages.Count - 1
                If Not DiagramHelper.Designer.ClosePage(False, True, True) Then
                    e.Cancel = True
                    Exit For
                End If
            Next

            MyBase.OnClosing(e)
        End Sub

        Private Function OpenCodeFile(filePath As String) As TextDocument
            Dim doc = ActiveDocument
            If doc IsNot Nothing Then
                If doc.IsNew AndAlso Not doc.IsDirty Then
                    CloseView(doc.MdiView)
                End If

            End If

            tabCode.IsSelected = True
            doc = New TextDocument(filePath)
            _DocumentTracker.TrackDocument(doc)

            Dim mdiView As New MdiView() With {
                .Document = doc,
                  .Width = viewsControl.ActualWidth - 100,
                  .Height = viewsControl.ActualHeight - 50
            }

            Canvas.SetTop(mdiView, 10)
            Canvas.SetLeft(mdiView, 10)
            mdiViews.Add(mdiView)
            mdiView.IsSelected = True
            doc.Focus(True)

            Return doc
        End Function

        Private Function SaveDocument(doc As TextDocument) As Boolean
            FormatCommand.Execute(Nothing, Me)

            If doc.IsNew Then
                Return SaveDocumentAs(doc)
            End If

            Try
                Dim designIsDirty = doc.PageKey = formDesigner.PageKey AndAlso doc.Form <> "" AndAlso formDesigner.HasChanges
                Dim genSaved = (doc.Form = "")

                If designIsDirty Then
                    SaveDesignInfo(doc)
                    genSaved = True
                End If

                If doc.IsDirty Then
                    doc.Save()
                    ' update the generated sb code to update event handlers added or removed
                    If Not genSaved Then IO.File.WriteAllText(doc.File & ".gen", doc.GenerateCodeBehind(formDesigner, False))
                End If

                Return True

            Catch ex As Exception
                Dim notificationButtons = Utility.MessageBox.Show(ResourceHelper.GetString("SaveFailed"), "Small Basic", String.Format(CultureInfo.CurrentCulture, ResourceHelper.GetString("SaveFailedReason"), New Object(0) {ex.Message}), Nf.No Or Nf.Yes, NotificationIcon.Information)

                If notificationButtons = Nf.Yes Then
                    Return SaveDocumentAs(doc)
                End If

                Return False
            End Try
        End Function

        Private Function SaveDocumentAs(document As TextDocument) As Boolean
            Dim saveFileDialog As New SaveFileDialog With {
                .Filter = ResourceHelper.GetString("SmallBasicFileFilter") & "|*.sb",
                .FileName = document.Form
            }

            If saveFileDialog.ShowDialog() Then
                Try
                    Dim fileName = saveFileDialog.FileName
                    Dim projectDir = IO.Path.GetDirectoryName(fileName)

                    If document.PageKey <> "" Then
                        Dim formName = document.Form.ToLower()
                        Dim newXamlfile = fileName.Replace(".sb", ".xaml")

                        For Each xamlFile In Directory.GetFiles(projectDir, "*.xaml")
                            If xamlFile.ToLower() = newXamlfile Then Continue For
                            If formName = DiagramHelper.Helper.GetFormNameFromXaml(xamlFile).ToLower() Then
                                MsgBox($"There is already a form named `{formName}` in this folder. Change the form name and try again, or save this file to another directory.")
                                Return False
                            End If
                        Next
                    End If

                    If File.Exists(fileName) Then File.Delete(fileName)
                    document.SaveAs(fileName)

                    If document.PageKey <> "" Then
                        Dim newFormName = Path.GetFileNameWithoutExtension(saveFileDialog.FileName)
                        Dim newDir = Path.GetDirectoryName(saveFileDialog.FileName)
                        Dim newFilePath = Path.Combine(newDir, newFormName)

                        fileName = newFilePath & ".sb.gen"
                        If File.Exists(fileName) Then File.Delete(fileName)

                        fileName = newFilePath & ".xaml"
                        If File.Exists(fileName) Then File.Delete(fileName)
                        Dim oldDir = IO.Path.GetDirectoryName(formDesigner.CodeFile)
                        formDesigner.XamlFile = fileName
                        formDesigner.CodeFile = saveFileDialog.FileName
                        formDesigner.HasChanges = True ' Force saving
                        SaveDesignInfo(document)
                        CopyImages(oldDir, newDir)
                    End If

                    ProjExplorer.ProjectDirectory = projectDir
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
                document.Focus()
                closePopHelp.After(1)

                Select Case Utility.MessageBox.Show(ResourceHelper.GetString("SaveDocumentBeforeClosing"), ResourceHelper.GetString("Title"), document.Title & ResourceHelper.GetString("DocumentModified"), Nf.Cancel Or Nf.No Or Nf.Yes, NotificationIcon.Information)
                    Case Nf.Yes
                        Return SaveDocument(document)
                    Case Nf.No
                        Return True
                    Case Else
                        Return False ' cancel
                End Select
            End If

            Return True
        End Function

        Private Sub RunProgram(Optional buildOnly As Boolean = False)
            Mouse.OverrideCursor = Cursors.Wait

            Dim doc As TextDocument

            If tabDesigner.IsSelected Then
                ' User can hit F5 on the designer.
                ' We need to save changes and generate the code behind
                doc = SaveDesignInfo()
                tabDesigner.IsSelected = True
            Else
                doc = ActiveDocument
            End If

            Format(doc)

            Dim code As String
            Dim genCode As String
            Dim errors As List(Of [Error])
            Dim filePath = doc.File

            Dim inputDir = If(filePath = "", "", Path.GetDirectoryName(filePath))
            Dim outputFileName = sVB.GetOutputFileName(
                filePath,
                doc.Form = "" AndAlso Not doc.IsTheGlobalFile
            )
            Dim formNames = doc.GetFormNames(True)

            doc.Errors.Clear()
            Dim parsers As List(Of Parser)

            Try
                If formNames.Count = 0 Then  ' Classic SB file without a form
                    parsers = New List(Of Parser)
                    code = doc.Text
                    'sVB.Compiler.ExeFile = outputFileName
                    If Not sVB.Compile("", code, doc, parsers) Then
                        Mouse.OverrideCursor = Nothing
                        Return
                    End If

                Else
                    parsers = sVB.CompileGlobalModule(inputDir, outputFileName, formNames, False)
                    If parsers Is Nothing Then
                        ' global file has errors
                        Mouse.OverrideCursor = Nothing
                        Return
                    End If

                    If filePath = "" Then
                        If doc.PageKey = "" Then
                            genCode = doc.GetCodeBehind(True)
                        Else
                            genCode = doc.GenerateCodeBehind(formDesigner, False)
                        End If

                        code = doc.Text
                        If Not sVB.Compile(genCode, code, doc, parsers) Then
                            Mouse.OverrideCursor = Nothing
                            Return
                        End If

                    Else
                        Dim currentFormKey = formDesigner.PageKey
                        Dim binDir = Path.GetDirectoryName(outputFileName)

                        For Each xamlFile In Directory.GetFiles(inputDir, "*.xaml")
                            Dim fName = DiagramHelper.Helper.GetFormNameFromXaml(xamlFile)
                            If fName = "" Then Continue For

                            If DiagramHelper.Designer.SavePageIfDirty(xamlFile) Then
                                Dim f2 = Path.Combine(binDir, Path.GetFileName(xamlFile))
                                Try
                                    File.Copy(xamlFile, f2, True)
                                Catch
                                End Try
                            End If

                            Dim sbCodeFile = xamlFile.Substring(0, xamlFile.Length - 4) + "sb"
                            Dim gen = viewsControl.SaveDocIfDirty(sbCodeFile)
                            If gen <> "" Then
                                genCode = gen
                            Else
                                Dim genCodefile = xamlFile.Substring(0, xamlFile.Length - 4) + "sb.gen"
                                If File.Exists(genCodefile) Then
                                    genCode = File.ReadAllText(genCodefile)
                                Else
                                    genCode = ""
                                End If
                            End If

                            If File.Exists(sbCodeFile) Then
                                code = File.ReadAllText(sbCodeFile)
                            Else
                                code = ""
                            End If

                            errors = sVB.Compile(genCode, code, False, False, formNames)

                            If errors.Count = 0 Then
                                Dim parser = sVB.Compiler.Parser
                                parser.IsMainForm = (fName = doc.Form)
                                parser.ClassName = "_SmallVisualBasic_" & fName.ToLower()
                                parsers.Add(parser)

                            Else
                                doc = OpenDocIfNot(sbCodeFile)
                                Call New RunAction(
                                    Sub()
                                        doc.ShowErrors(errors)
                                        tabCode.IsSelected = True
                                    End Sub
                                ).After(20)

                                Call New RunAction(
                                    Sub() doc.EditorControl.TextView.Caret.EnsureVisible()
                                ).After(500)

                                Mouse.OverrideCursor = Nothing
                                Return
                            End If
                        Next

                        DiagramHelper.Designer.SwitchTo(currentFormKey)
                    End If

                    If parsers.Count = 0 Then
                        If Not sVB.Compile("", doc.Text, doc, parsers) Then
                            Mouse.OverrideCursor = Nothing
                            Return
                        End If
                    End If
                End If

                sVB.Compiler.Build(parsers, outputFileName)

            Catch ex As Exception
                If errors Is Nothing Then errors = New List(Of [Error])
                errors.Add(New [Error](-1, 0, 0, ex.Message))
            End Try

            If errors?.Count > 0 Then
                doc.ShowErrors(errors)
                tabCode.IsSelected = True
                Mouse.OverrideCursor = Nothing
                Return
            End If

            Mouse.OverrideCursor = Nothing
            If buildOnly Then Return

            currentProgramProcess = Process.Start(outputFileName)
            currentProgramProcess.EnableRaisingEvents = True

            AddHandler currentProgramProcess.Exited,
                    Sub()
                        Dispatcher.BeginInvoke(DispatcherPriority.Normal, CType(
                             Function()
                                 Me.programRunningOverlay.Visibility = Visibility.Hidden
                                 currentProgramProcess = Nothing
                                 doc.Focus()
                                 Return True
                             End Function,
                             DispatcherOperationCallback), Nothing)
                    End Sub

            Me.processRunningMessage.Text = String.Format(ResourceHelper.GetString("ProgramRunning"), doc.Title)
            Me.programRunningOverlay.Visibility = Visibility.Visible
            Me.endProgramButton.Focus()

        End Sub



        Private Sub PublishDocument(document As TextDocument)
            If document.Text.Trim.Length < 50 Then
                MsgBox("Can't publish this program because its length is less than 50 chars.")
                Return
            End If

            RunProgram(True)

            If ActiveDocument.Errors.Count > 0 Then
                MsgBox("Please fix code errors before publishing this program.")
                Return
            End If

            Try
                Cursor = Cursors.Wait
                Dim service As New Service()
                Dim result = service.SaveProgram("", document.Text, document.BaseId)

                If Equals(result, "error") Then
                    Utility.MessageBox.Show(ResourceHelper.GetString("FailedToPublishToWeb"), ResourceHelper.GetString("Title"), ResourceHelper.GetString("PublishToWebFailedReason"), NotificationButtons.Close, NotificationIcon.Error)
                Else
                    Dim publishProgramDialog As New PublishProgramDialog(result)
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
                Dim stackPanel As New StackPanel()
                stackPanel.Orientation = Orientation.Vertical
                Dim stackPanel2 = stackPanel
                Dim textBlock As New TextBlock With {
                    .Text = ResourceHelper.GetString("ImportLocationOfProgramOnWeb"),
                    .Margin = New Thickness(0.0, 0.0, 4.0, 4.0)
                }
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
                messageBox.NotificationButtons = Nf.Cancel Or Nf.OK
                messageBox.NotificationIcon = NotificationIcon.Information
                textBox.Focus()

                If messageBox.Display() = Nf.OK Then
                    Dim service As New Service()
                    Dim baseId As String = textBox.Text.Trim()
                    Dim code = service.LoadProgram(baseId)

                    If Equals(code, "error") Then
                        Utility.MessageBox.Show(ResourceHelper.GetString("FailedToImportFromWeb"), ResourceHelper.GetString("Title"), ResourceHelper.GetString("ImportFromWebFailedReason"), NotificationButtons.Close, NotificationIcon.Error)
                    Else
                        code = code.Replace(vbLf, vbCrLf)
                        Dim newDocument As New TextDocument(Nothing)
                        newDocument.ContentType = "text.smallbasic"
                        newDocument.BaseId = baseId
                        newDocument.TextBuffer.Insert(0, code)
                        Dim service2 As New Service()
                        AddHandler service2.GetProgramDetailsCompleted,
                                Sub(o, e)
                                    Dim result = e.Result
                                    result.Category = ResourceHelper.GetString("Category" & result.Category)
                                    newDocument.ProgramDetails = result
                                End Sub

                        service2.GetProgramDetailsAsync(baseId)
                        _DocumentTracker.TrackDocument(newDocument)
                        Dim mdiView As New MdiView()
                        mdiView.Document = newDocument
                        mdiViews.Add(mdiView)
                        newDocument.Focus(True)
                    End If
                End If

            Catch ex As Exception
                Utility.MessageBox.Show(ResourceHelper.GetString("FailedToImportFromWeb"), ResourceHelper.GetString("Title"), String.Format(CultureInfo.CurrentUICulture, ResourceHelper.GetString("ReasonForFailure"), New Object(0) {ex.Message}), NotificationButtons.Close, NotificationIcon.Error)
            Finally
                Cursor = Cursors.Arrow
            End Try
        End Sub


        Private Sub OnCheckVersion(state As Object)
            Thread.Sleep(20000)

            Try
                Dim service As New Service()
                Dim currentVersion = service.GetCurrentVersion()
                Dim version = New Version(currentVersion)
                Dim version2 = Assembly.GetExecutingAssembly().GetName().Version

                'If version.CompareTo(version2) > 0 Then
                '    Dispatcher.BeginInvoke(CType(Sub() Me.updateAvailable.Visibility = Visibility.Visible, Action))
                'End If

            Catch
            End Try
        End Sub

        Private Sub OnClickNewVersionAvailable(sender As Object, e As RoutedEventArgs)
            Process.Start("http://smallbasic.com/download.aspx")
        End Sub

        Function OpenDocIfNot(filePath As String) As TextDocument
            If filePath = "" Then Return Nothing

            Dim docPath = Path.GetFullPath(filePath)
            For Each view As MdiView In Me.viewsControl.Items
                If view.Document.File.ToLower() = docPath.ToLower() Then
                    viewsControl.ChangeSelection(view)
                    Return view.Document
                End If
            Next

            Dim doc = OpenCodeFile(docPath)
            doc.IsNew = tempProjectPath <> "" AndAlso docPath.ToLower().StartsWith(tempProjectPath)
            Return doc
        End Function

        Function GetDocIfOpened() As TextDocument
            For Each view As MdiView In Me.viewsControl.Items
                If view.Document.PageKey = formDesigner.PageKey Then
                    Return view.Document
                End If
            Next
            Return Nothing
        End Function

        Public Function GetDocIfOpened(filePath As String) As TextDocument
            filePath = filePath.ToLower()
            For Each view As MdiView In Me.viewsControl.Items
                If view.Document.File.ToLower() = filePath Then
                    Return view.Document
                End If
            Next
            Return Nothing
        End Function

        Dim saveInfo As New RunAction(
                Sub()
                    SaveDesignInfo()
                    Me.ActiveDocument?.Focus()
                End Sub)

        Private Sub TabCode_Selected(sender As Object, e As RoutedEventArgs)
            If CStr(txtControlName.Tag) <> "" Then
                Dim controlIndex = CInt(txtControlName.Tag)
                If txtControlName.Text <> formDesigner.GetControlName(controlIndex) Then Return
            End If

            If formDesigner.Name = "global" Then
                Dim globFile = formDesigner.CodeFile
                If Not File.Exists(globFile) Then
                    File.Create(globFile).Close()
                End If
                Dim doc = OpenDocIfNot(globFile)
                doc.PageKey = formDesigner.PageKey
                doc?.Focus()

            ElseIf DiagramHelper.Designer.CurrentPage IsNot Nothing Then
                saveInfo.After(20)
            End If

            ' Note this prop isn't changed yet
            tabDesigner.IsSelected = False
            UpdateTitle()
        End Sub

        Private Sub TabDesigner_Selected(sender As Object, e As RoutedEventArgs)
            UpdateTitle()
            formDesigner.IsReloading = True
        End Sub

        Dim tempProjectPath As String

        Private Function GetProjectPath() As String
            If tempProjectPath <> "" Then Return tempProjectPath

            Dim projectName = Now.ToString("yy-MM-dd-HH-mm-ss", New CultureInfo("en-US"))
            tempProjectPath = Path.Combine(App.SvbUnsavedFolder, projectName).ToLower()
            DiagramHelper.Designer.TempProjectPath = tempProjectPath
            Return tempProjectPath
        End Function

        Private Function SaveDesignInfo(
                           Optional doc As TextDocument = Nothing,
                           Optional openDoc As Boolean = True,
                           Optional saveAs As Boolean = False
                    ) As TextDocument

            If DoNotOpenDefaultDoc Then
                DoNotOpenDefaultDoc = False
                Return Nothing
            End If

            Dim formName As String
            Dim xamlPath As String
            Dim formPath As String

            If CStr(txtControlName.Tag) <> "" Then
                TxtControlName_LostFocus(Nothing, Nothing)
            ElseIf CStr(txtControlText.Tag) <> "" Then
                TxtControlText_LostFocus(Nothing, Nothing)
            End If

            OpeningDoc = True

            If openDoc AndAlso formDesigner.CodeFile <> "" AndAlso
                        Not formDesigner.HasChanges Then
                doc = OpenDocIfNot(formDesigner.CodeFile)

                ' User may change the form name
                doc.Form = formDesigner.Name

                If doc.PageKey = "" Then doc.PageKey = formDesigner.PageKey
                OpeningDoc = False
                Return doc
            End If

            If formDesigner.XamlFile = "" Then
                If formDesigner.CodeFile = "" Then
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
                    If doc Is Nothing Then doc = GetDoc(formDesigner.CodeFile, openDoc)
                    formName = formDesigner.Name
                    doc.Form = formName

                    'xamlPath = Path.GetDirectoryName(formDesigner.CodeFilePath)
                    formPath = formDesigner.CodeFile.Substring(0, formDesigner.CodeFile.Length - 3)
                End If

                formDesigner.Name = formName
                formDesigner.DoSave(formPath & ".xaml")
                IO.File.Create(formPath & ".sb").Close()

            Else
                formPath = formDesigner.XamlFile.Substring(0, formDesigner.XamlFile.Length - 5)

                If doc Is Nothing Then doc = GetDoc(formPath & ".sb", openDoc)
                formName = formDesigner.Name
                doc.Form = formName

                If formDesigner.HasChanges OrElse saveAs Then formDesigner.DoSave()

            End If

            If doc Is Nothing Then doc = GetDoc(formPath & ".sb", openDoc)
            doc.Form = formName
            doc.PageKey = formDesigner.PageKey
            formDesigner.CodeFile = doc.File

            IO.File.WriteAllText(formPath & ".sb.gen", doc.GenerateCodeBehind(formDesigner, True))
            OpeningDoc = False
            Return doc
        End Function

        Function GetDoc(codeFilePath As String, openDoc As Boolean) As TextDocument
            If openDoc Then
                Return OpenDocIfNot(codeFilePath)
            Else
                Dim doc = GetDocIfOpened()
                If doc IsNot Nothing Then Return doc
                If Not File.Exists(codeFilePath) Then File.Create(codeFilePath).Close()
                Return New TextDocument(codeFilePath)
            End If

        End Function

        Private Sub MainWindow_Closed(sender As Object, e As EventArgs) Handles Me.Closed
            DeleteTempProjects(App.SvbTempFolder)
            DeleteTempProjects(App.SvbUnsavedFolder)
        End Sub

        Sub DeleteTempProjects(tempDir As String)
            For Each directory In IO.Directory.GetDirectories(tempDir)
                Try
                    Dim d = Date.ParseExact(
                            IO.Path.GetFileNameWithoutExtension(directory),
                            "yy-MM-dd-HH-mm-ss",
                            New CultureInfo("en-US")
                    )

                    If (Date.Now - d).TotalDays > 10 Then
                        Global.My.Computer.FileSystem.DeleteDirectory(directory, FileIO.DeleteDirectoryOption.DeleteAllContents)
                    End If

                Catch
                End Try
            Next
        End Sub

        Private Sub AddEventDefaultHandler(controlName As String)
            tabCode.IsSelected = True
            Mouse.OverrideCursor = Cursors.Wait

            Dim AddEventHandler As New RunAction(
                Sub()
                    If formDesigner.CodeFile <> "" Then
                        Dim doc = OpenDocIfNot(formDesigner.CodeFile)
                        If doc.ControlsInfo Is Nothing Then SaveDesignInfo(doc, False)

                        Dim key = controlName.ToLower()
                        Dim type = If(doc.ControlsInfo.ContainsKey(key), doc.ControlsInfo(key), "")

                        If doc.AddEventHandler(controlName, sb.PreCompiler.GetDefaultEvent(type), False) Then
                            doc.PageKey = formDesigner.PageKey
                            ' The code behind is saved before the new Handler is added.
                            ' We must make the designer dirty, to force saving this chamge in 
                            ' a next call to SaveDesignInfo()
                            formDesigner.HasChanges = True
                        End If
                    End If

                    tabCode.IsSelected = True
                    Mouse.OverrideCursor = Nothing
                End Sub)

            AddEventHandler.After(1)
        End Sub

        Private Sub FormDesigner_DiagramDoubleClick(control As UIElement)
            AddEventDefaultHandler(formDesigner.GetControlNameOrDefault(control))
        End Sub

        Private Sub FormDesigner_DoubleClick(sender As Object, e As MouseButtonEventArgs)
            e.Handled = True
            AddEventDefaultHandler(formDesigner.Name)
        End Sub


        Dim FirstTime As Boolean = True
        Dim SelectCodeTab As Boolean

        Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)
            If Not FirstTime Then Return
            FirstTime = False

            DiagramHelper.Designer.PagesGrid = DesignerGrid
            Dim HeaderPanelGrid As Grid = CType(VisualTreeHelper.GetChild(sVBTabs, 0), Grid)
            CType(Me.Content, Grid).Children.Remove(stkInfo)
            HeaderPanelGrid.Children.Add(stkInfo)
            stkInfo.SetValue(Panel.ZIndexProperty, 1000)
            UpdateTitle()

            AddHandler DiagramHelper.Designer.PageShown, AddressOf FormDesigner_CurrentPageChanged

            'DiagramHelper.Designer.SetDefaultPropertiesSub =
            '    Sub()
            '        DiagramHelper.Designer.SetDefaultProperties()
            '        'Set any defaults you want here

            '    End Sub

        End Sub

        Dim formNameChanged As Boolean

        Function SavePage(oldPath As String, saveAs As Boolean) As Boolean
            Dim msg = DiagramHelper.Helper.FormNameExists(formDesigner)
            If msg <> "" Then
                MsgBox(msg)
                If saveAs Then formDesigner.XamlFile = oldPath
                Return False
            End If

            Dim newFormName = Path.GetFileNameWithoutExtension(formDesigner.XamlFile)
            Dim newDir = Path.GetDirectoryName(formDesigner.XamlFile)
            Dim newFilePath = Path.Combine(newDir, newFormName)
            Dim fileName = newFilePath & ".sb"

            Dim oldCodeFile = formDesigner.CodeFile
            If oldCodeFile = "" AndAlso oldPath <> "" Then oldCodeFile = oldPath.Substring(0, oldPath.Length - 4) & "sb"
            SaveDesignInfo(Nothing, False, saveAs)

            If saveAs Then
                Dim doc = GetDocIfOpened()
                Try
                    If oldCodeFile?.ToLower() = fileName.ToLower() Then
                        doc?.Save()
                    Else

                        If oldCodeFile <> "" AndAlso File.Exists(fileName) Then
                            If IO.File.ReadAllText(fileName) = "" OrElse MsgBox(
                                    $"There is a file with the same name `{fileName}`. Do you want to overwrite if? ",
                                    MsgBoxStyle.YesNo Or MsgBoxStyle.Exclamation) = MsgBoxResult.Yes Then
                                File.Delete(fileName)
                            End If
                        End If

                        If doc IsNot Nothing Then
                            doc.SaveAs(fileName)
                        ElseIf File.Exists(oldCodeFile) Then
                            File.Copy(oldCodeFile, fileName, True)
                        End If

                        formDesigner.CodeFile = fileName
                    End If

                Catch ex As Exception
                    MsgBox(ex.Message)
                    Return False
                End Try

                If oldCodeFile <> "" Then
                    CopyImages(IO.Path.GetDirectoryName(oldCodeFile), newDir)
                End If
            End If

            ProjExplorer.ProjectDirectory = formDesigner.XamlFile
            Return True
        End Function

        Private Sub CopyImages(oldDir As String, newDir As String)

            For Each f In IO.Directory.EnumerateFiles(oldDir)
                Select Case IO.Path.GetExtension(f).ToLower().TrimStart("."c)
                    Case "bmp", "jpg", "jpeg", "png", "gif", "ico"
                        Dim f2 = IO.Path.Combine(newDir, IO.Path.GetFileName(f))
                        Try
                            IO.File.Copy(f, f2, True)
                        Catch
                        End Try
                End Select
            Next
        End Sub

        Private Sub FormDesigner_CurrentPageChanged(index As Integer)
            If index <0 AndAlso index > DiagramHelper.Designer.GlobalFileIndex Then
                UpdateTitle()
                UpdateTextBoxes()
                txtControlName.IsEnabled = True
                txtControlText.IsEnabled = True
                Return
            End If

            formDesigner = DiagramHelper.Designer.CurrentPage
            ZoomBox.Designer = formDesigner
            OFExplorer.Designer = formDesigner
            ToolBox.Designer = formDesigner

            If index = DiagramHelper.Designer.GlobalFileIndex Then
                txtControlName.IsEnabled = False
                txtControlText.IsEnabled = False
                ProjExplorer.SelectedGlobalFile()
            Else
                txtControlName.IsEnabled = True
                txtControlText.IsEnabled = True
                ProjExplorer.ProjectDirectory = formDesigner.XamlFile
            End If

            ' Remove the handler if exists not to be called twice
            RemoveHandler formDesigner.DiagramDoubleClick, AddressOf FormDesigner_DiagramDoubleClick
            AddHandler formDesigner.DiagramDoubleClick, AddressOf FormDesigner_DiagramDoubleClick

            RemoveHandler formDesigner.MouseDoubleClick, AddressOf FormDesigner_DoubleClick
            AddHandler formDesigner.MouseDoubleClick, AddressOf FormDesigner_DoubleClick

            RemoveHandler formDesigner.SelectionChanged, AddressOf FormDesigner_SelectionChanged
            AddHandler formDesigner.SelectionChanged, AddressOf FormDesigner_SelectionChanged

            formDesigner.SavePage = AddressOf SavePage

            UpdateTitle()

            If index > -1 Then
                OFExplorer.SelectedIndex = index
            End If

            UpdateTextBoxes()

        End Sub

        Dim FocusTxtName As New DiagramHelper.RunAction(20,
                 Sub()
                     txtControlName.Focus()
                     txtControlName.SelectAll()
                 End Sub)

        Dim FocusTxtText As New DiagramHelper.RunAction(20,
                 Sub() txtControlText.SelectAll())

        Dim ExitSelectionChanged As Boolean

        Private Sub FormDesigner_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
            If ExitSelectionChanged Then Return

            Dim controlIndex As Integer
            If Not formDesigner.ItemDeleted Then
                If CStr(txtControlName.Tag) <> "" Then
                    controlIndex = CInt(txtControlName.Tag)
                    If formDesigner.Items.Count > controlIndex AndAlso txtControlName.Text <> formDesigner.GetControlName(controlIndex) Then
                        If Not CommitName() Then
                            ' Re-select the control. this event can fire b4 lostfocus event of the textbox
                            ExitSelectionChanged = True
                            If controlIndex = -1 Then
                                formDesigner.SelectedItem = Nothing
                            Else
                                ' Note: setting selectedIndex doesn't work!!
                                formDesigner.SelectedItem = formDesigner.Items(controlIndex)
                            End If
                            ExitSelectionChanged = False

                            ' goback to the textbox
                            FocusTxtName.Start()
                            Return
                        End If
                    End If
                End If

                If CStr(txtControlText.Tag) <> "" Then
                    controlIndex = CInt(txtControlText.Tag)
                    formDesigner.SetControlText(controlIndex, txtControlText.Text)
                End If
            End If

            UpdateTextBoxes()

        End Sub

        Sub UpdateTextBoxes()
            txtControlName.Tag = ""
            txtControlText.Tag = ""

            Dim controlIndex = formDesigner.SelectedIndex
            If formDesigner.SelectedIndex = -1 Then
                txtControlName.Text = formDesigner.Name
                txtControlText.Text = formDesigner.Text
            Else
                txtControlName.Text = formDesigner.GetControlName()
                txtControlText.Text = formDesigner.GetControlText()
            End If
        End Sub

        Private Sub UpdateTitle()
            If tabDesigner.IsSelected Then
                StkToolbar.Visibility = Visibility.Visible
                grdNameText.Visibility = Visibility.Visible
                txtTitle.Text = "Design - "
                txtForm.Text = If(formDesigner.XamlFile = "",
                        formDesigner.Name & " *",
                        Path.GetFileName(formDesigner.XamlFile)
                )

            Else
                StkToolbar.Visibility = Visibility.Collapsed
                grdNameText.Visibility = Visibility.Collapsed
                txtTitle.Text = "Code - "
                txtForm.Text = If(formDesigner.CodeFile = "",
                        $"{formDesigner.Name}.sb",
                        Path.GetFileName(formDesigner.CodeFile)
                )
            End If
        End Sub

        Dim OpeningDoc As Boolean
        Friend Shared FilesToOpen As New List(Of String)

        Private Sub ViewsControl_ActiveDocumentChanged()
            If OpeningDoc Then Return

            Dim currentView = viewsControl.SelectedItem
            CmbOpenedDocs.SelectedItem = currentView
            If currentView Is Nothing Then Return

            Dim doc = currentView.Document
            Dim codeFilePath = doc.File

            If doc.PageKey = "" Then
                If codeFilePath <> "" Then
                    If Path.GetFileNameWithoutExtension(codeFilePath).ToLower() = "global" Then
                        doc.PageKey = DiagramHelper.Designer.SwitchTo(DiagramHelper.Helper.GlobalFileName)
                        formDesigner.CodeFile = codeFilePath
                        ProjExplorer.ProjectDirectory = codeFilePath
                        Return
                    End If

                    Dim pagePath = codeFilePath.Substring(0, codeFilePath.Length - 3) & ".xaml"
                    If File.Exists(pagePath) Then
                        doc.PageKey = DiagramHelper.Designer.SwitchTo(pagePath)
                        formDesigner.CodeFile = codeFilePath
                    Else
                        ' Do nothing to allow opening old sb files without a form
                        ' If you want to attach a form, comment the next line,
                        ' ' and  umcomment the 2 lineslines after.
                        doc.PageKey = ""
                    End If
                End If

            Else
                DiagramHelper.Designer.SwitchTo(doc.PageKey)
                formDesigner.CodeFile = codeFilePath
                If Path.GetFileNameWithoutExtension(codeFilePath) = "global" Then
                    ProjExplorer.ProjectDirectory = codeFilePath
                End If
            End If

        End Sub

        Private Sub OFExplorer_ItemDoubleClick(sender As Object, e As MouseButtonEventArgs)
            tabCode.IsSelected = True
        End Sub

        Private Sub CloseButton_Click(sender As Object, e As RoutedEventArgs)
            If tabCode.IsSelected Then
                If ActiveDocument IsNot Nothing Then CloseView(ActiveDocument.MdiView)
            Else
                DiagramHelper.Designer.ClosePage()
            End If

        End Sub

        Function CommitName() As Boolean
            If CStr(txtControlName.Tag) = "" Then Return True

            Dim controlIndex = CInt(txtControlName.Tag)
            Dim newName = txtControlName.Text.Trim()
            If newName = "" Then
                Beep()
                ExitSelectionChanged = True
                tabCode.IsSelected = False
                tabDesigner.IsSelected = True
                ExitSelectionChanged = False
                Return False
            End If

            Dim oldName = formDesigner.GetControlName(controlIndex)
            If oldName = newName Then Return True

            newName = newName(0).ToString().ToUpper & If(newName.Length > 1, newName.Substring(1), "")
            txtControlName.Text = newName

            If IsKeyword(newName) Then
                Return False
            End If

            If controlIndex < 0 Then
                Dim msg = DiagramHelper.Helper.FormNameExists(formDesigner, newName)
                If msg <> "" Then
                    ExitSelectionChanged = True
                    tabCode.IsSelected = False
                    tabDesigner.IsSelected = True
                    MsgBox(msg)
                    ExitSelectionChanged = False
                    Return False
                End If
            End If

            If Not formDesigner.SetControlName(controlIndex, newName) Then
                ExitSelectionChanged=True
                tabCode.IsSelected = False
                tabDesigner.IsSelected = True
                ExitSelectionChanged = False
                Return False
            End If

            Dim doc = OpenDocIfNot(formDesigner.CodeFile)
            If doc IsNot Nothing Then
                doc.FixEventHandlers(oldName, newName)
                SaveDesignInfo(doc)
            End If

            Return True
        End Function

        Private Function IsKeyword(newName As String) As Boolean
            Dim name = newName.ToLower()
            Dim msg = $"'{newName}' is an sVB keyword and can't be used as a name. you can add a control prefix to the name, such as `Frm{newName}` or `Txt{newName}`."

            If name = "me" OrElse name = "global" Then
                ExitSelectionChanged = True
                tabCode.IsSelected = False
                tabDesigner.IsSelected = True
                MsgBox(msg)
                ExitSelectionChanged = False
                Return True
            End If

            Dim tokens = LineScanner.GetTokens(name, 0)
            If tokens.Count > 1 Then
                ExitSelectionChanged = True
                tabCode.IsSelected = False
                tabDesigner.IsSelected = True
                MsgBox("Form and control names can't start with a ni,ber nor contain spaces or any symbols. Use `_` instead.")
                ExitSelectionChanged = False
                Return True
            End If

            Select Case tokens(0).ParseType
                Case ParseType.Keyword, ParseType.Operator
                    ExitSelectionChanged = True
                    tabCode.IsSelected = False
                    tabDesigner.IsSelected = True
                    MsgBox(msg)
                    ExitSelectionChanged = False
                    Return True
                Case Else
                    Return False
            End Select

        End Function

        Private Sub TxtControlName_PreviewKeyDown(sender As Object, e As KeyEventArgs)
            Select Case e.Key
                Case Key.Enter
                    CommitName()
                    e.Handled = True

                Case Key.Escape
                    UpdateTextBoxes()
                    e.Handled = True

                Case Key.Space
                    Beep()
                    e.Handled = True
            End Select
        End Sub

        Private Sub TxtControlName_LostFocus(sender As Object, e As RoutedEventArgs)
            txtControlName.SelectionLength = 0
            If FocusTxtName.Started Then Return

            If CommitName() Then
                txtControlName.Tag = ""
            Else
                If e IsNot Nothing Then e.Handled = True
                FocusTxtName.Start()
            End If
        End Sub

        Private Sub TxtControlName_PreviewTextInput(sender As Object, e As TextCompositionEventArgs)
            Select Case e.Text.ToLower
                Case "a" To "z", "_", "0" To "9"
                    ' allowed
                Case Else
                    VisualBasic.Beep()
                    e.Handled = True
            End Select
        End Sub

        Private Sub TxtControlText_PreviewKeyDown(sender As Object, e As KeyEventArgs)
            Select Case e.Key
                Case Key.Enter
                    If CStr(txtControlText.Tag) <> "" Then
                        Dim controlIndex = CInt(txtControlText.Tag)
                        formDesigner.SetControlText(controlIndex, txtControlText.Text)
                    End If
                    e.Handled = True

                Case Key.Escape
                    UpdateTextBoxes()
                    e.Handled = True
            End Select

        End Sub

        Private Sub TxtControlText_LostFocus(sender As Object, e As RoutedEventArgs)
            txtControlText.SelectionLength = 0
            If CStr(txtControlText.Tag) = "" Then Return

            Dim controlIndex = CInt(txtControlText.Tag)
            formDesigner.SetControlText(controlIndex, txtControlText.Text)
            txtControlText.Tag = ""
        End Sub

        Private Sub TxtControl_GotFocus(sender As Object, e As RoutedEventArgs)
            Dim txt As TextBox = sender
            txt.Tag = formDesigner.SelectedIndex
        End Sub

        Dim DoNotOpenDefaultDoc As Boolean

        Private Sub Window_ContentRendered(sender As Object, e As EventArgs)
            Dim doc As TextDocument
            Dim closePage = False

            If FilesToOpen.Count > 0 Then
                For Each fileName In FilesToOpen
                    fileName = fileName.ToLower()
                    If Path.GetExtension(fileName) = ".xaml" OrElse (
                                Path.GetExtension(fileName) = ".sb" AndAlso
                                File.Exists(fileName.Substring(0, fileName.Length - 2) & "xaml")
                            ) Then
                        closePage = True
                        Exit For
                    End If
                Next

                If closePage Then
                    DiagramHelper.Designer.ClosePage(False, True)
                End If
            End If

            For Each fileName In FilesToOpen
                Dim file = fileName.ToLower()
                If file.EndsWith(".sb") Then
                    doc = OpenDocIfNot(fileName)
                    SelectCodeTab = True

                ElseIf file.EndsWith(".xaml") Then
                    DiagramHelper.Designer.SwitchTo(fileName)
                    SelectCodeTab = False
                End If
            Next

            ' Load the code editor
            If doc Is Nothing Then
                tabCode.IsSelected = True
            Else
                DoNotOpenDefaultDoc = Not closePage
            End If

            If SelectCodeTab Then
                If doc IsNot Nothing Then
                    tabCode.IsSelected = True
                    Call New RunAction(
                        Sub()
                            viewsControl.ChangeSelection(doc.MdiView)
                        End Sub
                    ).After(10)
                End If

            Else
                Dim selectDesigner As New RunAction(
                    Sub()
                        tabDesigner.IsSelected = True
                        ' Crate a new doc and close it, to consume the first time load delay.
                        doc = New TextDocument(Nothing)
                        Dim mdiView As New MdiView() With {.Document = doc}
                        mdiViews.Add(mdiView)
                        CloseView(mdiView)
                    End Sub)
                selectDesigner.After(10)
            End If
        End Sub

        Private Sub MainWindow_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
            CompletionProvider.popHelp = PopHelp
        End Sub

        Private Sub MainWindow_PreviewMouseDown(sender As Object, e As MouseButtonEventArgs) Handles Me.PreviewMouseDown
            If GetParent(Of MainWindow)(e.OriginalSource) IsNot Nothing Then
                PopHelp.IsOpen = False
            End If
        End Sub

        Private Sub MainWindow_PreviewKeyDown(sender As Object, e As KeyEventArgs) Handles Me.PreviewKeyDown
            If e.Key = Key.Escape Then
                PopHelp.IsOpen = False
            End If
        End Sub

        Dim closePopHelp As New RunAction(Sub() PopHelp.IsOpen = False)
        Private Sub PopHelp_Opened(sender As Object, e As EventArgs)
            closePopHelp.After(10000)
        End Sub

        Private Sub TxtControlName_GotFocus(sender As Object, e As RoutedEventArgs) Handles txtControlName.GotFocus
            FocusTxtName.Start()
        End Sub

        Private Sub TxtControlText_GotFocus(sender As Object, e As RoutedEventArgs) Handles txtControlText.GotFocus
            FocusTxtText.Start()
        End Sub

        Private Sub MainWindow_Deactivated(sender As Object, e As EventArgs) Handles Me.Deactivated
            PopHelp.IsOpen = False
        End Sub

        Private Sub ProjExplorer_FileNameChanged(oldFileName As String, newFileName As String)
            formDesigner.CodeFile = newFileName
            formDesigner.XamlFile = newFileName.Substring(0, newFileName.Length - 2) & "xaml"
            Dim doc = GetDoc(oldFileName, False)
            Dim genFile = newFileName & ".gen"
            doc.File = newFileName
            File.WriteAllText(genFile, doc.GenerateCodeBehind(formDesigner, True))
        End Sub

        Dim ShowGlobalFile As New RunAction(
                Sub()
                    tabCode.IsSelected = True
                    tabDesigner.IsSelected = False
                End Sub
        )

        Private Sub tabDesigner_PreviewMouseDoubleClick(sender As Object, e As MouseButtonEventArgs)
            If e.OriginalSource IsNot DesignerGrid Then Return
            If Not formDesigner.IsEnabled Then
                ShowGlobalFile.After(1)
            End If
        End Sub

        Private Sub tabDesigner_PreviewKeyDown(sender As Object, e As KeyEventArgs)
            If e.Key = Key.F5 Then
                RunProgram()
                e.Handled = True
            End If
        End Sub

        Private Sub CmbOpenedDocs_GotFocus(sender As Object, e As RoutedEventArgs)
            CmbOpenedDocs.IsDropDownOpen = True
        End Sub

        Private Sub CmbOpenedDocs_GotFocus(sender As Object, e As MouseButtonEventArgs)

        End Sub

        Private Sub CmbOpenedDocs_PreviewMouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs)
            If e.OriginalSource.GetType().Name = "TextBoxView" Then
                CmbOpenedDocs.IsDropDownOpen = True
                e.Handled = True
            End If
        End Sub

        Private Sub CmbOpenedDocs_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
            Dim view = CType(CmbOpenedDocs.SelectedItem, MdiView)
            view?.Document.Focus()
        End Sub

        Private Sub ProjExplorer_FileDeleted()
            Dim doc = GetDocIfOpened()
            If doc IsNot Nothing Then
                doc.IsDirty = False
                CloseView(doc.MdiView)
            End If

            formDesigner.HasChanges = False
            DiagramHelper.Designer.ClosePage(True, True)
        End Sub

        Private Sub BtnSavePage_Click(sender As Object, e As RoutedEventArgs)
            formDesigner.Save()
            formDesigner.Focus()
        End Sub

        Private Sub BtnNewPage_Click(sender As Object, e As RoutedEventArgs)
            DiagramHelper.Designer.OpenNewPage()
            formDesigner.Focus()
        End Sub

        Private Sub BtnOpenPage_Click(sender As Object, e As RoutedEventArgs)
            DiagramHelper.Designer.Open()
            formDesigner.Focus()
        End Sub

        Private Sub BtnRun_Click(sender As Object, e As RoutedEventArgs)
            RunProgram()
        End Sub

    End Class


End Namespace
