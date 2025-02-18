﻿Imports Microsoft.Nautilus.Text
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
Imports Microsoft.SmallVisualBasic.Library
Imports DiagramHelper
Imports MS.Internal.Xaml

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
            "Run",
            "RunProgramCommand",
            GetType(MainWindow)
        )

        Public Shared DebugCommand As New RoutedUICommand(
            "Debug",
            "DebugCommand",
            GetType(MainWindow)
         )

        Public Shared ToggleBreakpointCommand As New RoutedUICommand(
            "Breakpoint",
            "ToggleBreakpointCommand",
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

        Public Shared ShortStepOverCommand As New RoutedUICommand(
            "Block Over",
            "ShortStepOverCommand",
            GetType(MainWindow)
        )

        Public Shared StepIntoCommand As New RoutedUICommand(
            "Step Into",
            "StepIntoCommand",
            GetType(MainWindow)
        )

        Public Shared StepOutCommand As New RoutedUICommand(
            "Step Out",
            "StepOutCommand",
            GetType(MainWindow)
        )

        Public Shared ShortStepOutCommand As New RoutedUICommand(
            "Block Out",
            "ShortStepOutCommand",
            GetType(MainWindow)
        )

        Public Shared PauseCommand As New RoutedUICommand(
            "Pause",
            "PauseCommand",
            GetType(MainWindow)
        )

        Public Shared BreakpointCommand As New RoutedUICommand(
            ResourceHelper.GetString("BreakpointCommand"),
            ResourceHelper.GetString("BreakpointCommand"),
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
                SaveSetting("SmallVisualBasic", "Files", "Open", IO.Path.GetDirectoryName(openFileDialog.FileName))
                OpenDocIfNot(openFileDialog.FileName).Focus()
            End If

        End Sub

        Private Sub CanFileSave(sender As Object, e As CanExecuteRoutedEventArgs)
            e.CanExecute = tabCode.IsSelected AndAlso ActiveDocument IsNot Nothing
        End Sub

        Private Sub OnFileSave(sender As Object, e As RoutedEventArgs)
            SaveDocument(ActiveDocument)
        End Sub

        Private Sub OnFileSaveAs(sender As Object, e As RoutedEventArgs)
            SaveDocumentAs(ActiveDocument)
        End Sub

        Private Sub CanEditCut(sender As Object, e As CanExecuteRoutedEventArgs)
            e.CanExecute = tabCode.IsSelected AndAlso ActiveDocument IsNot Nothing AndAlso Not ActiveDocument.EditorControl.TextView.Selection.IsEmpty
        End Sub

        Private Sub OnEditCut(sender As Object, e As RoutedEventArgs)
            Dim doc = ActiveDocument
            doc.EditorControl.EditorOperations.CutSelection(doc.UndoHistory)
        End Sub

        Private Sub CanEditCopy(sender As Object, e As CanExecuteRoutedEventArgs)
            e.CanExecute = tabCode.IsSelected AndAlso ActiveDocument IsNot Nothing AndAlso Not ActiveDocument.EditorControl.TextView.Selection.IsEmpty
        End Sub


        Private Sub OnEditCopy(sender As Object, e As RoutedEventArgs)
            ActiveDocument.EditorControl.EditorOperations.CopySelection()
        End Sub

        Private Sub CanEditPaste(sender As Object, e As CanExecuteRoutedEventArgs)
            e.CanExecute = tabCode.IsSelected AndAlso ActiveDocument IsNot Nothing AndAlso ActiveDocument.EditorControl.EditorOperations.CanPaste
        End Sub

        Private Sub OnEditPaste(sender As Object, e As RoutedEventArgs)
            Dim doc = ActiveDocument
            doc.EditorControl.EditorOperations.Paste(doc.UndoHistory)
        End Sub

        Private Sub CanEditUndo(sender As Object, e As CanExecuteRoutedEventArgs)
            e.CanExecute = tabCode.IsSelected AndAlso ActiveDocument IsNot Nothing AndAlso ActiveDocument.UndoHistory.CanUndo
        End Sub

        Private Sub OnEditUndo(sender As Object, e As RoutedEventArgs)
            ActiveDocument.UndoHistory.Undo(1)
        End Sub

        Private Sub CanEditRedo(sender As Object, e As CanExecuteRoutedEventArgs)
            e.CanExecute = tabCode.IsSelected AndAlso ActiveDocument IsNot Nothing AndAlso ActiveDocument.UndoHistory.CanRedo
        End Sub

        Private Sub OnEditRedo(sender As Object, e As RoutedEventArgs)
            ActiveDocument.UndoHistory.Redo(1)
        End Sub

        Private Sub OnCloseItem(sender As Object, e As RequestCloseEventArgs)
            Dim mdiView = TryCast(e.Item, MdiView)
            CloseView(mdiView)
        End Sub

        Friend Sub EvaluateExpression(text As String)
            Dim debugger = GetDebugger(True)
            If debugger IsNot Nothing Then
                Dim result = debugger.EvaluateExpression(text)
                If result.HasValue Then
                    HelpPanel.ShowExpression(text, result.Value)
                End If
            End If
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
            e.CanExecute = tabCode.IsSelected AndAlso ActiveDocument IsNot Nothing AndAlso ActiveDocument.Text.Trim().Length > 0
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
            Dim debugger = GetDebugger(True)
            e.CanExecute = tabDesigner.IsSelected OrElse
                     ActiveDocument IsNot Nothing OrElse
                     (debugger IsNot Nothing AndAlso debugger.IsActive)
        End Sub

        Private Sub OnProgramRun(sender As Object, e As RoutedEventArgs)
            RunProgram()
        End Sub

        Private Sub OnProgramEnd(sender As Object, e As RoutedEventArgs)
            Dim debugger = GetDebugger(True)
            If debugger?.IsActive Then
                debugger.End()
            ElseIf currentProgramProcess IsNot Nothing AndAlso Not currentProgramProcess.HasExited Then
                currentProgramProcess.Kill()
                currentProgramProcess = Nothing
                Me.programRunningOverlay.Visibility = Visibility.Hidden
            End If
        End Sub

        Private Sub OnStepOver(sender As Object, e As RoutedEventArgs)
            GetDebugger(False).StepOver()
        End Sub

        Private Sub OnShortStepOver(sender As Object, e As RoutedEventArgs)
            GetDebugger(False).ShortStepOver()
        End Sub

        Private Sub OnStepInto(sender As Object, e As RoutedEventArgs)
            GetDebugger(False).StepInto()
        End Sub

        Private Sub OnStepOut(sender As Object, e As RoutedEventArgs)
            GetDebugger(False).StepOut()
        End Sub

        Private Sub OnShortStepOut(sender As Object, e As RoutedEventArgs)
            GetDebugger(False).ShortStepOut()
        End Sub

        Private Sub OnPause(sender As Object, e As RoutedEventArgs)
            GetDebugger(False).Pause()
        End Sub

        Public Sub DisplayDebugCommands(show As Boolean)
            Dim v1, v2 As Visibility
            If show Then
                v1 = Visibility.Visible
                v2 = Visibility.Collapsed
            Else
                v1 = Visibility.Collapsed
                v2 = Visibility.Visible
            End If

            Ribbon.SetVisible(RoutedRibbonCommand.Commands(StepOverCommand), v1)
            Ribbon.SetVisible(RoutedRibbonCommand.Commands(ShortStepOverCommand), v1)
            Ribbon.SetVisible(RoutedRibbonCommand.Commands(StepOutCommand), v1)
            Ribbon.SetVisible(RoutedRibbonCommand.Commands(ShortStepOutCommand), v1)
            Ribbon.SetVisible(RoutedRibbonCommand.Commands(PauseCommand), v1)
            Ribbon.SetVisible(RoutedRibbonCommand.Commands(EndProgramCommand), v1)

            Ribbon.SetVisible(RoutedRibbonCommand.Commands(NewCommand), v2)
            Ribbon.SetVisible(RoutedRibbonCommand.Commands(SaveCommand), v2)
            Ribbon.SetVisible(RoutedRibbonCommand.Commands(SaveAsCommand), v2)
            Ribbon.SetVisible(RoutedRibbonCommand.Commands(CutCommand), v2)
            Ribbon.SetVisible(RoutedRibbonCommand.Commands(PasteCommand), v2)
            Ribbon.SetVisible(RoutedRibbonCommand.Commands(UndoCommand), v2)
            Ribbon.SetVisible(RoutedRibbonCommand.Commands(RedoCommand), v2)
            Ribbon.SetVisible(RoutedRibbonCommand.Commands(DebugCommand), v2)
            Ribbon.SetVisible(RoutedRibbonCommand.Commands(ExportToVisualBasicCommand), v2)
        End Sub

        Private Sub CanToggleBreakpoint(sender As Object, e As CanExecuteRoutedEventArgs)
            Dim doc = ActiveDocument
            If Not tabCode.IsSelected OrElse doc Is Nothing Then
                e.CanExecute = False
                Return
            End If

            Dim pos = doc.EditorControl.TextView.Caret.Position.CharacterIndex
            If pos < 0 Then
                e.CanExecute = False
                Return
            End If

            Dim line = doc.TextBuffer.CurrentSnapshot.GetLineFromPosition(pos)
            If line.GetText().Trim() = "" Then
                e.CanExecute = False
                Return
            End If

            Dim lineNumber = line.LineNumber
            Dim span = doc.GetFullStatementSpan(lineNumber)

            If span Is Nothing Then
                e.CanExecute = False
            ElseIf doc.Breakpoints.Contains(lineNumber) Then
                e.CanExecute = True
            Else
                Dim token = LineScanner.GetFirstToken(line.GetText(), 0)
                e.CanExecute = token.Type <> TokenType.Comment
            End If
        End Sub

        Private Sub OnToggleBreakpoint(sender As Object, e As RoutedEventArgs)
            Dim doc = ActiveDocument
            Dim pos = doc.EditorControl.TextView.Caret.Position.CharacterIndex
            Dim line = doc.TextBuffer.CurrentSnapshot.GetLineFromPosition(pos)
            Dim lineNumber = line.LineNumber
            Dim showBreakpoint As Boolean
            doc.ToggleBreakpoint(lineNumber, showBreakpoint)
            doc.EditorControl.LineNumberMargin.DrawBreakpoint(lineNumber, showBreakpoint)
        End Sub

        Private Sub CanWebSave(sender As Object, e As CanExecuteRoutedEventArgs)
            e.CanExecute = tabCode.IsSelected AndAlso ActiveDocument IsNot Nothing
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
            If Me.programRunningOverlay.Visibility = Visibility.Visible OrElse
                      GetDebugger(True)?.IsActive Then
                OnProgramEnd(Nothing, Nothing)
                e.Cancel = True
                Return
            End If

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
        Private Function OpenCodeFile(filePath As String, Optional focus As Boolean = True) As TextDocument
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
            Me.UpdateLayout()
            dispatcher.BeginInvoke(
                Sub()
                    mdiView.IsSelected = True
                    If focus Then doc.Focus(True)
                End Sub)

            Return doc
        End Function

        Private Function SaveDocument(doc As TextDocument) As Boolean
            FormatCommand.Execute(Nothing, Me)

            If doc.IsNew Then
                Return SaveDocumentAs(doc)
            End If

            Try
                Dim designIsDirty = doc.PageKey = formDesigner.PageKey AndAlso doc.Form <> "" AndAlso formDesigner.MustSaveDesign
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
                    SaveSetting("SmallVisualBasic", "Files", "Open", projectDir)

                    If document.PageKey <> "" Then
                        Dim formName = document.Form.ToLower()
                        Dim newXamlfile = fileName.Replace(".sb", ".xaml").ToLower()

                        For Each xamlFile In Directory.GetFiles(projectDir, "*.xaml")
                            If xamlFile.ToLower() = newXamlfile Then Continue For
                            If formName = DiagramHelper.Helper.GetFormNameFromXaml(xamlFile).ToLower() Then
                                MsgBox($"There is already a form named `{formName}` in this folder. Change the form name and try again, or save this file to another directory.")
                                Return False
                            End If
                        Next
                    End If

                    If IO.File.Exists(fileName) Then IO.File.Delete(fileName)
                    document.SaveAs(fileName)

                    If document.PageKey <> "" Then
                        Dim newFormName = Path.GetFileNameWithoutExtension(saveFileDialog.FileName)
                        Dim newDir = Path.GetDirectoryName(saveFileDialog.FileName)
                        Dim newFilePath = Path.Combine(newDir, newFormName)

                        fileName = newFilePath & ".sb.gen"
                        If IO.File.Exists(fileName) Then IO.File.Delete(fileName)

                        fileName = newFilePath & ".xaml"
                        If IO.File.Exists(fileName) Then IO.File.Delete(fileName)
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
                document.Activate()
                PopHelp.IsOpen = False

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
            Dim debugger = GetDebugger(True)
            If debugger?.IsActive Then
                debugger.Continue()
            Else
                Input.Mouse.OverrideCursor = Cursors.Wait
                BuildAndRun(buildOnly)
                Input.Mouse.OverrideCursor = Nothing
            End If
        End Sub

        Friend Function GetDebugger(acceptNull As Boolean) As ProgramDebugger
            Dim key As String

            If tabDesigner.IsSelected Then
                key = ProjExplorer.ProjectDirectory
            Else
                Dim doc = ActiveDocument
                If doc IsNot Nothing AndAlso doc.Form = "" AndAlso Not doc.IsTheGlobalFile Then
                    ' Single code file. Use it as a key
                    key = doc.File
                Else ' A file within a project. Use the folder as a key
                    key = ProjExplorer.ProjectDirectory
                End If
            End If

            Return ProgramDebugger.GetDebugger(key, acceptNull)
        End Function

        Private Sub OnDebugProgram()
            GetDebugger(False).Run(False)
        End Sub


        Friend Function BuildAndRun(buildOnly As Boolean) As List(Of Parser)
            Dim doc As TextDocument

            If tabDesigner.IsSelected Then
                ' User can hit F5 on the designer.
                ' We need to save changes and generate the code behind
                Try
                    doc = SaveDesignInfo()
                Catch ex As Exception
                    MsgBox(ex.Message)
                    Return Nothing
                End Try
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
            SmallVisualBasic.Library.Program.AppExe = New Primitive(outputFileName)
            doc.Errors.Clear()
            Dim formNames = doc.GetFormNames()
            Library.Program.FormNames = formNames

            If doc.Errors.Count > 0 Then
                Dim formFile = doc.Errors(0)
                Dim formName = doc.Errors(1)
                doc.Errors.Clear()
                DiagramHelper.Designer.SwitchTo(formFile)
                dispatcher.BeginInvoke(
                    DispatcherPriority.Background,
                    Sub()
                        tabCode.IsSelected = False
                        tabDesigner.IsSelected = True
                        MsgBox($"There is another form in the project with the name `{formName}`. Each form must have a unique programmable name. Please fix this before running the program.")
                        Input.Mouse.OverrideCursor = Nothing
                        ProjExplorer.FilesList.Focus()
                    End Sub)
                Return Nothing
            End If

            Dim parsers As List(Of Parser)
            sVB.ClassName = ""

            Try
                If formNames.Count = 0 Then  ' Classic SB file without a form
                    parsers = New List(Of Parser)
                    code = doc.Text
                    'sVB.Compiler.ExeFile = outputFileName
                    If Not sVB.Compile("", code, doc, parsers) Then
                        Input.Mouse.OverrideCursor = Nothing
                        Return Nothing
                    End If

                Else
                    parsers = sVB.CompileGlobalModule(inputDir, outputFileName, formNames, False)
                    If parsers Is Nothing Then
                        ' global file has errors
                        Input.Mouse.OverrideCursor = Nothing
                        Return Nothing
                    End If

                    If filePath = "" Then
                        If doc.PageKey = "" Then
                            genCode = doc.GetCodeBehind(True)
                        Else
                            genCode = doc.GenerateCodeBehind(formDesigner, False)
                        End If

                        code = doc.Text
                        If Not sVB.Compile(genCode, code, doc, parsers) Then
                            Input.Mouse.OverrideCursor = Nothing
                            Return Nothing
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
                                    IO.File.Copy(xamlFile, f2, True)
                                Catch
                                End Try
                            End If

                            Dim sbCodeFile = xamlFile.Substring(0, xamlFile.Length - 4) + "sb"
                            Dim gen = viewsControl.SaveDocIfDirty(sbCodeFile)
                            If gen <> "" Then
                                genCode = gen
                            Else
                                Dim genCodefile = xamlFile.Substring(0, xamlFile.Length - 4) + "sb.gen"
                                If IO.File.Exists(genCodefile) Then
                                    genCode = IO.File.ReadAllText(genCodefile)
                                Else
                                    genCode = ""
                                End If
                            End If

                            If IO.File.Exists(sbCodeFile) Then
                                code = IO.File.ReadAllText(sbCodeFile)
                            Else
                                code = ""
                            End If

                            sVB.ClassName = WinForms.Forms.FormPrefix & fName.ToLower()
                            errors = sVB.Compile(genCode, code, False, False, formNames)

                            If errors.Count = 0 Then
                                Dim parser = sVB.Compiler.Parser
                                parser.DocPath = sbCodeFile
                                parser.IsMainForm = (fName = doc.Form)
                                parsers.Add(parser)

                            Else
                                doc = OpenDocIfNot(sbCodeFile)
                                dispatcher.BeginInvoke(
                                     System.Windows.Threading.DispatcherPriority.Background,
                                     Sub()
                                         doc.ShowErrors(errors)
                                         tabCode.IsSelected = True
                                     End Sub
                                )

                                dispatcher.BeginInvoke(
                                     System.Windows.Threading.DispatcherPriority.Background,
                                     Sub() doc.EditorControl.TextView.Caret.EnsureVisible()
                                )

                                Input.Mouse.OverrideCursor = Nothing
                                Return Nothing
                            End If
                        Next

                        DiagramHelper.Designer.SwitchTo(currentFormKey)
                    End If

                    If parsers.Count = 0 Then
                        If Not sVB.Compile("", doc.Text, doc, parsers) Then
                            Input.Mouse.OverrideCursor = Nothing
                            Return Nothing
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
                Input.Mouse.OverrideCursor = Nothing
                Return Nothing
            End If

            If buildOnly Then Return parsers

            currentProgramProcess = Process.Start(outputFileName)
            currentProgramProcess.EnableRaisingEvents = True

            AddHandler currentProgramProcess.Exited,
                    Sub()
                        dispatcher.BeginInvoke(DispatcherPriority.Normal, CType(
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

            Return parsers
        End Function

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

            If MsgBox(
                    "Are you sure you want to publish this program?",
                    MsgBoxStyle.YesNo Or MsgBoxStyle.Question Or MsgBoxStyle.DefaultButton2,
                    "Confirmation") = MsgBoxResult.No Then
                Return
            End If

            Try
                Cursor = Cursors.Wait
                Dim service As New Service()
                Dim result = service.SaveProgram("", document.Text, document.BaseId)

                If result = "error" Then
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
                Dim stackPanel As New StackPanel()
                stackPanel.Orientation = Orientation.Vertical
                Dim textBlock As New TextBlock With {
                    .Text = ResourceHelper.GetString("ImportLocationOfProgramOnWeb"),
                    .Margin = New Thickness(0.0, 0.0, 4.0, 4.0)
                }
                Dim textBox As New TextBox() With {
                    .FontSize = 32.0,
                    .FontWeight = FontWeights.Bold,
                    .FontFamily = New FontFamily("Courier New"),
                    .Foreground = Brushes.DimGray,
                    .Margin = New Thickness(0.0, 4.0, 4.0, 4.0),
                    .MinWidth = 300.0
                }
                stackPanel.Children.Add(textBlock)
                stackPanel.Children.Add(textBox)
                Dim messageBox As New Utility.MessageBox() With {
                    .Description = ResourceHelper.GetString("ImportFromWeb"),
                    .Title = ResourceHelper.GetString("Title"),
                    .OptionalContent = stackPanel,
                    .NotificationButtons = Nf.Cancel Or Nf.OK,
                    .NotificationIcon = NotificationIcon.Information
                }

                If messageBox.Display() = Nf.OK Then
                    Input.Mouse.OverrideCursor = Cursors.Wait
                    Dim service As New Service()
                    Dim baseId As String = textBox.Text.Trim()
                    Dim code = service.LoadProgram(baseId)

                    If Equals(code, "error") Then
                        Utility.MessageBox.Show(ResourceHelper.GetString("FailedToImportFromWeb"), ResourceHelper.GetString("Title"), ResourceHelper.GetString("ImportFromWebFailedReason"), NotificationButtons.Close, NotificationIcon.Error)
                    Else
                        code = code.Replace(vbLf, vbCrLf)
                        Dim newDoc As New TextDocument(Nothing)
                        _DocumentTracker.TrackDocument(newDoc)
                        newDoc.ContentType = "text.smallbasic"
                        newDoc.BaseId = baseId
                        newDoc.TextBuffer.Insert(0, code)
                        AddHandler service.GetProgramDetailsCompleted,
                                Sub(o, e)
                                    Dim result = e.Result
                                    result.Category = ResourceHelper.GetString("Category" & result.Category)
                                    newDoc.ProgramDetails = result
                                End Sub

                        service.GetProgramDetailsAsync(baseId)
                        Dim mdiView As New MdiView() With {
                            .Document = newDoc,
                            .Width = Me.Width - 100,
                            .Height = Me.Height - 350
                        }
                        mdiViews.Add(mdiView)
                        Me.UpdateLayout()
                        textBox.Focus()

                        MainWindow.dispatcher.BeginInvoke(
                            Sub()
                                newDoc.Focus(True)
                            End Sub)
                    End If
                End If

            Catch ex As Exception
                Utility.MessageBox.Show(ResourceHelper.GetString("FailedToImportFromWeb"), ResourceHelper.GetString("Title"), String.Format(CultureInfo.CurrentUICulture, ResourceHelper.GetString("ReasonForFailure"), New Object(0) {ex.Message}), NotificationButtons.Close, NotificationIcon.Error)
            Finally
                Input.Mouse.OverrideCursor = Nothing
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

        Function OpenDocIfNot(filePath As String, Optional focus As Boolean = True) As TextDocument
            If filePath = "" Then Return Nothing

            Dim doc As TextDocument
            Dim docPath = Path.GetFullPath(filePath).ToLower()
            For Each view As MdiView In Me.viewsControl.Items
                doc = view.Document
                If doc.File.ToLower() = docPath Then
                    viewsControl.ChangeSelection(view)
                    ProgramDebugger.LockRunningDoc(doc)
                    Return doc
                End If
            Next

            doc = OpenCodeFile(docPath, focus)
            doc.IsNew = (tempProjectPath <> "" AndAlso docPath.StartsWith(tempProjectPath))
            ProgramDebugger.LockRunningDoc(doc)
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

        Private Shared dispatcher As Dispatcher = Application.Current.Dispatcher

        Private Sub TabCode_Selected(sender As Object, e As RoutedEventArgs)
            If CStr(txtControlName.Tag) <> "" Then
                Dim controlIndex = CInt(txtControlName.Tag)
                If txtControlName.Text <> formDesigner.GetControlName(controlIndex) Then Return
            End If

            If formDesigner.Name = "global" Then
                Dim globFile = formDesigner.CodeFile
                If Not IO.File.Exists(globFile) Then
                    IO.File.Create(globFile).Close()
                End If
                Dim doc = OpenDocIfNot(globFile)
                doc.PageKey = formDesigner.PageKey
                doc?.Focus()

            ElseIf DiagramHelper.Designer.CurrentPage IsNot Nothing Then
                dispatcher.BeginInvoke(
                       DispatcherPriority.Background,
                       Sub()
                           Try
                               SaveDesignInfo()
                               Me.ActiveDocument?.Focus()
                           Catch ex As Exception
                               MsgBox(ex.Message)
                           End Try
                       End Sub)
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
        Dim CurrentProccessID As String = Process.GetCurrentProcess().Id.ToString()

        Private Function GetProjectPath() As String
            If tempProjectPath <> "" Then Return tempProjectPath

            Dim projectName = Now.ToString("yy-MM-dd-HH-mm-ss", New CultureInfo("en-US"))
            tempProjectPath = Path.Combine(App.SvbUnsavedFolder, projectName).ToLower()
            DiagramHelper.Designer.TempProjectPath = tempProjectPath
            SaveSetting("SmallVisualBasic", "Backup", "LastProject", tempProjectPath)
            SaveSetting("SmallVisualBasic", "Backup", "ProccessID", CurrentProccessID)
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
                        Not formDesigner.MustSaveDesign Then
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
                    IO.File.Create(formPath & ".sb").Close()
                Else
                    If doc Is Nothing Then doc = GetDoc(formDesigner.CodeFile, openDoc)
                    formName = formDesigner.Name
                    doc.Form = formName
                    formPath = formDesigner.CodeFile.Substring(0, formDesigner.CodeFile.Length - 3)
                End If

                formDesigner.Name = formName
                formDesigner.DoSave(formPath & ".xaml")
            Else
                formPath = formDesigner.XamlFile.Substring(0, formDesigner.XamlFile.Length - 5)

                If doc Is Nothing Then doc = GetDoc(formPath & ".sb", openDoc)
                formName = formDesigner.Name
                doc.Form = formName

                If formDesigner.MustSaveDesign OrElse saveAs Then
                    formDesigner.DoSave()
                End If
            End If

            If doc Is Nothing Then doc = GetDoc(formPath & ".sb", openDoc)
            doc.Form = formName
            doc.PageKey = formDesigner.PageKey
            formDesigner.CodeFile = doc.File

            Try
                IO.File.WriteAllText(formPath & ".sb.gen", doc.GenerateCodeBehind(formDesigner, True))
            Finally
                OpeningDoc = False
            End Try

            Return doc
        End Function

        Function GetDoc(codeFilePath As String, openDoc As Boolean) As TextDocument
            If openDoc Then
                Return OpenDocIfNot(codeFilePath)
            Else
                Dim doc = GetDocIfOpened()
                If doc IsNot Nothing Then Return doc
                If Not IO.File.Exists(codeFilePath) Then IO.File.Create(codeFilePath).Close()
                Return New TextDocument(codeFilePath)
            End If

        End Function

        Private Sub MainWindow_Closed(sender As Object, e As EventArgs) Handles Me.Closed
            DeleteTempProjects(App.SvbTempFolder)
            DeleteTempProjects(App.SvbUnsavedFolder)
            ' Window is closed normally
            SaveSetting("SmallVisualBasic", "Backup", "LastProject", "")
            SaveSetting("SmallVisualBasic", "Backup", "ProccessID", "")
            SaveSetting("SmallVisualBasic", "User", "ScaleFactor", TextDocument.ScaleFactor.ToString())
            SmallVisualBasic.Library.Program.End()
            SmallVisualBasic.Library.TextWindow.Close()
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

        Private Sub AddEventDefaultHandler(controlName As String, Optional eventName As String = "")
            tabCode.IsSelected = True
            Input.Mouse.OverrideCursor = Cursors.Wait

            dispatcher.BeginInvoke(
                   System.Windows.Threading.DispatcherPriority.Background,
                   Sub()
                       Try
                           If formDesigner.CodeFile <> "" Then
                               Dim doc = OpenDocIfNot(formDesigner.CodeFile)
                               If doc.ControlsInfo Is Nothing Then SaveDesignInfo(doc, False)

                               Dim key = controlName.ToLower()
                               Dim type = If(doc.ControlsInfo.ContainsKey(key), doc.ControlsInfo(key), "")
                               If eventName = "" Then eventName = sb.PreCompiler.GetDefaultEvent(type)

                               If doc.AddEventHandler(controlName, eventName, False) Then
                                   doc.PageKey = formDesigner.PageKey
                                   ' The code behind is saved before the new Handler is added.
                                   ' We must make the designer dirty, to force saving this chamge in 
                                   ' a next call to SaveDesignInfo()
                                   formDesigner.HasChanges = True
                               End If
                           End If

                           tabCode.IsSelected = True
                           Input.Mouse.OverrideCursor = Nothing
                       Catch ex As Exception
                           MsgBox(ex.Message)
                       End Try
                   End Sub)
        End Sub

        Private Sub FormDesigner_DiagramDoubleClick(control As UIElement)
            AddEventDefaultHandler(formDesigner.GetControlNameOrDefault(control))
        End Sub

        Private Sub FormDesigner_DoubleClick(sender As Object, e As MouseButtonEventArgs)
            If TypeOf e.OriginalSource IsNot Canvas Then Return

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
            AddHandler DiagramHelper.Designer.OnMenuItemClicked, AddressOf FormDesigner_OnMenuItemClicked
            AddHandler DiagramHelper.Designer.OnOpeningCodeFile, AddressOf Designer_OnOpeningCodeFile
        End Sub

        Private Sub Designer_OnOpeningCodeFile(fileName As String)
            tabCode.IsSelected = True
            tabDesigner.IsSelected = False
            dispatcher.BeginInvoke(Sub() OpenDocIfNot(fileName).Focus(), DispatcherPriority.Background)
        End Sub

        Private Sub FormDesigner_OnMenuItemClicked(sender As MenuItem)
            AddEventDefaultHandler(sender.Name, If(sender.Items.Count = 0, "", "OnOpen"))
        End Sub

        Dim formNameChanged As Boolean

        Function SavePage(oldPath As String, saveAs As Boolean) As Boolean
            Try
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

                            If oldCodeFile <> "" AndAlso IO.File.Exists(fileName) Then
                                If IO.File.ReadAllText(fileName) = "" OrElse MsgBox(
                                    $"There Is a file with the same name `{fileName}`. Do you want to overwrite if? ",
                                    MsgBoxStyle.YesNo Or MsgBoxStyle.Exclamation) = MsgBoxResult.Yes Then
                                    IO.File.Delete(fileName)
                                End If
                            End If

                            If doc IsNot Nothing Then
                                doc.SaveAs(fileName)
                            ElseIf IO.File.Exists(oldCodeFile) Then
                                IO.File.Copy(oldCodeFile, fileName, True)
                            End If

                            formDesigner.CodeFile = fileName
                        End If

                    Catch ex As Exception
                        MsgBox(ex.Message)
                        Return False
                    End Try

                    If oldCodeFile <> "" Then
                        CopyImages(IO.Path.GetDirectoryName(oldCodeFile), newDir)
                        If oldCodeFile.ToLower.StartsWith(tempProjectPath) Then
                            Try
                                Directory.Delete(Path.GetDirectoryName(oldCodeFile), True)
                            Catch
                            End Try
                        End If
                    End If
                End If

                ProjExplorer.ProjectDirectory = formDesigner.XamlFile
                Return True

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            Return False
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
            If index < 0 AndAlso index > DiagramHelper.Designer.GlobalFileIndex Then
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

        Dim focusTxtNameStarted As Boolean = False

        Private Sub FocusTxtName()
            focusTxtNameStarted = True
            dispatcher.BeginInvoke(
                   DispatcherPriority.Background,
                   Sub()
                       focusTxtNameStarted = False
                       txtControlName.Focus()
                       txtControlName.SelectAll()
                   End Sub)
        End Sub

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
                            FocusTxtName()
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
                        tabDesigner.IsSelected = True
                        dispatcher.BeginInvoke(DispatcherPriority.Background,
                             Sub()
                                 doc.PageKey = DiagramHelper.Designer.SwitchTo(DiagramHelper.Helper.GlobalFileName)
                                 formDesigner.CodeFile = codeFilePath
                                 ProjExplorer.ProjectDirectory = codeFilePath
                                 tabDesigner.IsSelected = False
                                 tabCode.IsSelected = True
                             End Sub)
                        Return
                    End If

                    Dim pagePath = codeFilePath.Substring(0, codeFilePath.Length - 3) & ".xaml"
                    If IO.File.Exists(pagePath) Then
                        tabDesigner.IsSelected = True
                        dispatcher.BeginInvoke(DispatcherPriority.Background,
                             Sub()
                                 doc.PageKey = DiagramHelper.Designer.SwitchTo(pagePath)
                                 formDesigner.CodeFile = codeFilePath
                                 tabDesigner.IsSelected = False
                                 tabCode.IsSelected = True
                             End Sub)
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
                ExitSelectionChanged = True
                tabCode.IsSelected = False
                tabDesigner.IsSelected = True
                ExitSelectionChanged = False
                Return False
            End If

            Dim doc = OpenDocIfNot(formDesigner.CodeFile)
            If doc IsNot Nothing Then
                doc.FixEventHandlers(oldName, newName)
                Try
                    SaveDesignInfo(doc)
                Catch ex As Exception
                    MsgBox(ex.Message)
                    Return False
                End Try
            End If

            Return True
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
            If focusTxtNameStarted Then Return

            If CommitName() Then
                txtControlName.Tag = ""
            Else
                If e IsNot Nothing Then e.Handled = True
                FocusTxtName()
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
                                IO.File.Exists(fileName.Substring(0, fileName.Length - 2) & "xaml")
                            ) Then
                        closePage = True
                        Exit For
                    End If
                Next

                If closePage Then
                    DiagramHelper.Designer.ClosePage(False, True)
                End If

            Else
                tempProjectPath = GetSetting("SmallVisualBasic", "Backup", "LastProject", "")
                If tempProjectPath <> "" AndAlso Directory.Exists(tempProjectPath) Then
                    Dim proccessID = GetSetting("SmallVisualBasic", "Backup", "ProccessID")
                    If proccessID = "" Then
                        LoadedUnsavedProject()
                    Else
                        Dim sVBProcessName = Process.GetCurrentProcess().ProcessName
                        Dim id = CInt(proccessID)
                        Dim isRunning = False

                        For Each p In Process.GetProcesses()
                            If p.Id = id AndAlso p.ProcessName = sVBProcessName Then
                                isRunning = True
                                Exit For
                            End If
                        Next
                        If Not isRunning Then LoadedUnsavedProject()
                    End If
                End If
            End If

            tabDesigner.IsSelected = False
            tabCode.IsSelected = True

            dispatcher.BeginInvoke(
                  System.Windows.Threading.DispatcherPriority.Background,
                  Sub()
                      For Each fileName In FilesToOpen
                          Dim file = fileName.ToLower()
                          If file.EndsWith(".sb") Then
                              doc = OpenDocIfNot(fileName)
                              SelectCodeTab = True

                          ElseIf file.EndsWith(".xaml") Then
                              dispatcher.BeginInvoke(DispatcherPriority.Background,
                                    Sub()
                                        tabCode.IsSelected = False
                                        tabDesigner.IsSelected = True
                                        DiagramHelper.Designer.SwitchTo(fileName)
                                    End Sub
                              )
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
                              dispatcher.BeginInvoke(
                                     System.Windows.Threading.DispatcherPriority.Background,
                                     Sub()
                                         viewsControl.ChangeSelection(doc.MdiView)
                                         DoNotOpenDefaultDoc = False
                                     End Sub
                              )
                          End If

                      Else
                          dispatcher.BeginInvoke(
                              System.Windows.Threading.DispatcherPriority.Background,
                              Sub()
                                  tabDesigner.IsSelected = True
                                  ' Crate a new doc and close it, to consume the first time load delay.
                                  doc = New TextDocument(Nothing)
                                  Dim mdiView As New MdiView() With {.Document = doc}
                                  mdiViews.Add(mdiView)
                                  CloseView(mdiView)
                                  DoNotOpenDefaultDoc = False
                              End Sub)
                      End If
                  End Sub)
        End Sub

        Private Sub LoadedUnsavedProject()
            DiagramHelper.Designer.TempProjectPath = tempProjectPath
            DiagramHelper.Designer.ClosePage(False, True, False)
            DiagramHelper.Designer.TempKeyNum = 0

            For Each dir1 In Directory.GetDirectories(tempProjectPath)
                Dim xamlFiles = Directory.GetFiles(dir1, "*.xaml")
                If xamlFiles.Length > 0 Then
                    Dim f = xamlFiles(0)
                    DiagramHelper.Designer.SwitchTo(f)
                    DoNotOpenDefaultDoc = True
                    FilesToOpen.Add(f.Substring(0, f.Length - 4) + "sb")
                End If
            Next
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
            Select Case e.Key
                Case Key.Escape
                    If PopHelp.IsOpen Then
                        PopHelp.IsOpen = False
                    Else
                        PopError.IsOpen = False
                    End If
                Case Key.F4
                    If tabDesigner.IsSelected Then formDesigner.ShowProperties()
            End Select
        End Sub

        Dim closePopHelp As New DispatcherTimer(
                 New TimeSpan(0, 0, 10),
                 DispatcherPriority.Background,
                 Sub() PopHelp.IsOpen = False,
                 dispatcher
        )

        Private Sub PopHelp_Opened(sender As Object, e As EventArgs)
            closePopHelp.IsEnabled = True
        End Sub

        Private Sub PopHelp_Closed(sender As Object, e As EventArgs)
            closePopHelp.IsEnabled = False
            HelpPanel.DontShowHelp = False
        End Sub

        Private Sub TxtControlName_GotFocus(sender As Object, e As RoutedEventArgs) Handles txtControlName.GotFocus
            FocusTxtName()
        End Sub

        Private Sub TxtControlText_GotFocus(sender As Object, e As RoutedEventArgs) Handles txtControlText.GotFocus
            dispatcher.BeginInvoke(
                DispatcherPriority.Background,
                Sub() txtControlText.SelectAll()
            )
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
            IO.File.WriteAllText(genFile, doc.GenerateCodeBehind(formDesigner, True))
        End Sub

        Private Sub tabDesigner_PreviewMouseDoubleClick(sender As Object, e As MouseButtonEventArgs)
            If e.OriginalSource IsNot DesignerGrid Then Return
            If Not formDesigner.IsEnabled Then
                dispatcher.BeginInvoke(
                    DispatcherPriority.Background,
                    Sub()
                        tabCode.IsSelected = True
                        tabDesigner.IsSelected = False
                    End Sub)

            End If
        End Sub

        Private Sub tabDesigner_PreviewKeyDown(sender As Object, e As KeyEventArgs)
            If e.Key = Key.F5 Then
                If Input.Keyboard.Modifiers = ModifierKeys.Control Then
                    OnDebugProgram()
                Else
                    RunProgram()
                End If
                e.Handled = True
            End If
        End Sub

        Private Sub CmbOpenedDocs_GotFocus(sender As Object, e As RoutedEventArgs)
            CmbOpenedDocs.IsDropDownOpen = True
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
            SaveDesignInfo(Nothing, False)
        End Sub

        Private Sub BtnOpenPage_Click(sender As Object, e As RoutedEventArgs)
            DiagramHelper.Designer.Open()
        End Sub

        Private Sub BtnRun_Click(sender As Object, e As RoutedEventArgs)
            RunProgram()
        End Sub

        Private Sub BtnDebug_Click(sender As Object, e As RoutedEventArgs)
            OnDebugProgram()
        End Sub

        Private Sub BtnProps_Click(sender As Object, e As RoutedEventArgs)
            formDesigner.ShowProperties()
        End Sub
    End Class
End Namespace
