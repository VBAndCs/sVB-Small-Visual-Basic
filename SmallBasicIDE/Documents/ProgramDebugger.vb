Imports Microsoft.Nautilus.Text
Imports Microsoft.Nautilus.Text.Editor
Imports Microsoft.SmallVisualBasic.Engine
Imports Microsoft.SmallVisualBasic.LanguageService
Imports Microsoft.Windows.Controls
Imports System
Imports System.Collections.Generic
Imports System.Windows.Media
Imports System.Windows.Threading

Namespace Microsoft.SmallVisualBasic.Documents
    Public Class ProgramDebugger


        Private Shared debuggers As New Dictionary(Of String, ProgramDebugger)

        Private Sub New()

        End Sub

        Public Shared Function GetDebugger(project As String, acceptNull As Boolean) As ProgramDebugger
            ' Only one active debugger is allowed, so return it if found.
            ' This will prevent running a new project while debigging another
            For Each debugger In debuggers.Values
                If debugger.IsActive Then Return debugger
            Next

            Dim key = project.ToLower()
            If debuggers.ContainsKey(key) Then
                Return debuggers(key)
            ElseIf acceptNull Then
                Return Nothing
            Else
                Dim debugger As New ProgramDebugger()
                debuggers(key) = debugger
                Return debugger
            End If
        End Function

        Public ReadOnly Property Dispatcher As Dispatcher
            Get
                Return Document.EditorControl.Dispatcher
            End Get
        End Property

        Public ReadOnly Property Document As TextDocument
            Get
                Dim docPath = ProgramEngine.CurrentRunner.Parser.DocPath
                Dim m = Helper.MainWindow
                Dim doc = If(docPath = "", m.ActiveDocument, m.OpenDocIfNot(docPath, False))
                Return doc
            End Get
        End Property

        Public Property IsActive As Boolean

        Public Property ProgramEngine As ProgramEngine

        Private Function CreateEngine() As ProgramEngine
            If _ProgramEngine Is Nothing Then
                Dim parsers = Helper.MainWindow.BuildAndRun(True)
                If parsers Is Nothing OrElse parsers.Count = 0 Then Return Nothing

                _ProgramEngine = New Engine.ProgramEngine(parsers)
                AddHandler Library.Program.ProgramTerminated, AddressOf OnProgramTerminated
                AddHandler _ProgramEngine.DebuggerStateChanged, AddressOf OnCurrentStateChanged
            End If
            Return _ProgramEngine
        End Function


        Public ReadOnly Property TextBuffer As ITextBuffer
            Get
                Return Document.TextBuffer
            End Get
        End Property

        Dim isEvaluating As Boolean
        Friend Function Evaluate(item As Completion.CompletionItem) As Library.Primitive?
            If isEvaluating Then Return Nothing

            Dim doc = Helper.MainWindow.ActiveDocument
            Dim runner = _ProgramEngine.GetEvaluationRunner(doc.File)
            If runner Is Nothing Then Return Nothing

            isEvaluating = True
            Dim result? As Library.Primitive = Nothing

            If item.ObjectName Is Nothing OrElse
                        item.ItemType = Completion.CompletionItemType.Control OrElse
                        item.ItemType = Completion.CompletionItemType.LocalVariable OrElse
                        item.ItemType = Completion.CompletionItemType.GlobalVariable Then

                Dim key = runner.GetKey(item.DefinitionIdintifier)
                If key = "" Then
                    key = item.DisplayName.ToLower()
                End If
                result = GetValue(key, runner.Fields)

            ElseIf item.ObjectName.ToLower() = "global" Then
                Dim key = item.DisplayName.ToLower()
                result = GetValue(key, _ProgramEngine.GlobalRunner.Fields)

            ElseIf item.ItemType = Completion.CompletionItemType.DynamicPropertyName Then
                Dim subName = doc.GetCurrentSubName()
                Dim arrKey = If(subName = "", item.ObjectName, $"{subName}.{item.ObjectName}").ToLower()
                Dim arr = GetValue(arrKey, runner.Fields)

                If arr.HasValue Then
                    result = Library.Primitive.GetArrayValue(arr.Value, item.DisplayName)

                ElseIf subName <> "" Then
                    arrKey = item.ObjectName.ToLower()
                    arr = GetValue(arrKey, runner.Fields)
                    If arr.HasValue Then
                        result = Library.Primitive.GetArrayValue(arr.Value, item.DisplayName)
                    End If
                End If

            Else
                Dim key = item.DisplayName.ToLower()
                Dim objectName As New Token() With {.Text = item.ObjectName, .Type = TokenType.Identifier}
                Dim type = runner.Parser.SymbolTable.GetTypeInfo(objectName)

                Dim propInfo = TryCast(item.MemberInfo, Reflection.PropertyInfo)
                If propInfo IsNot Nothing Then
                    result = CType(propInfo.GetValue(Nothing, Nothing), Library.Primitive)
                Else
                    Dim name = GetValue(objectName.LCaseText, runner.Fields)
                    If name Is Nothing Then
                        Dim subName = doc.GetCurrentSubName()
                        If subName = "" Then
                            isEvaluating = False
                            Return Nothing
                        End If

                        Dim objKey = $"{subName}.{item.ObjectName}".ToLower()
                        name = GetValue(objKey, runner.Fields)
                    End If

                    If name IsNot Nothing Then
                        Dim methodName = "get" & key
                        If Not type.Methods.ContainsKey(methodName) Then
                            type = runner.TypeInfoBag.Types("control")
                        End If
                        Dim methodInfo = type.Methods(methodName)
                        result = CType(methodInfo.Invoke(Nothing, New Object() {name}), Library.Primitive)
                    End If
                End If
            End If

            isEvaluating = False
            Return result
        End Function

        Private Function GetValue(key As String, fields As Dictionary(Of String, Library.Primitive)) As Library.Primitive?
            If fields.ContainsKey(key) Then
                Return fields(key)
            Else
                Return Nothing
            End If
        End Function

        Public Sub Pause()
            If IsActive Then _ProgramEngine.Pause()
        End Sub

        Public Sub [End]()
            If IsActive Then Library.Program.End()
        End Sub

        Public Sub Run(breakOnStart As Boolean)
            If IsActive Then
                _ProgramEngine.Continue()
            Else
                Dim m = Helper.MainWindow
                m.tabCode.IsSelected = True
                m.tabDesigner.IsEnabled = False
                m.DisplayDebugCommands(True)
                m.Dispatcher.BeginInvoke(
                    DispatcherPriority.Background,
                    Sub()
                        _ProgramEngine = CreateEngine()
                        If _ProgramEngine IsNot Nothing Then
                            _ProgramEngine.RunProgram(breakOnStart)
                            For Each parser In _ProgramEngine.Parsers
                                Dim doc = m.GetDocIfOpened(parser.DocPath)
                                If doc IsNot Nothing Then
                                    If _ProgramEngine.Runners.ContainsKey(parser) Then
                                        _ProgramEngine.Runners(parser).Breakpoints = doc.Breakpoints
                                    ElseIf doc.Breakpoints.Count > 0 Then
                                        Dim r = New ProgramRunner(_ProgramEngine, parser)
                                        r.Breakpoints = doc.Breakpoints
                                        _ProgramEngine.Runners(parser) = r
                                    End If
                                    doc.ReadOnly = True
                                End If
                            Next
                            _IsActive = True
                        End If
                    End Sub)
            End If
        End Sub

        Dim lastDoc As TextDocument

        Private Sub OnCurrentStateChanged(runner As ProgramRunner)
            Dim m = Helper.MainWindow
            m.Dispatcher.Invoke(
                Sub()
                    Dim docPath = runner.Parser.DocPath.ToLower()
                    Dim CurDoc = If(
                            docPath = "" OrElse docPath = m.ActiveDocument.File.ToLower(),
                            m.ActiveDocument,
                            m.OpenDocIfNot(docPath, False)
                    )
                    Dim offset = runner.DocLineOffset
                    Dim lineNumber = runner.CurrentLineNumber - offset
                    Dim currentState = runner.DebuggerState
                    Dim markLine = (currentState = DebuggerState.Paused OrElse currentState = DebuggerState.Error)
                    Dim doc = If(lastDoc, CurDoc)
                    Dim stMarkerProvider = doc.StatementMarkerProvider
                    Dim marker = doc.ExecutionMarker

                    If marker IsNot Nothing Then
                        doc.ExecutionMarker = Nothing
                        If doc.Breakpoints.Contains(marker.LineNumber) Then
                            marker.MarkerColor = Colors.Red
                        Else
                            stMarkerProvider.RemoveMarker(marker)
                        End If
                    End If

                    If _ProgramEngine Is Nothing Then Return ' Program ended

                    If lastDoc IsNot Nothing Then
                        doc = CurDoc
                        stMarkerProvider = doc.StatementMarkerProvider
                    End If

                    doc.ReadOnly = True

                    If markLine Then
                        ActivateIDE()
                        Dim line = doc.TextBuffer.CurrentSnapshot.GetLineFromLineNumber(lineNumber)

                        If doc.Breakpoints.Contains(lineNumber) Then
                            marker = stMarkerProvider.GetMarker(lineNumber)
                            If marker IsNot Nothing Then
                                marker.MarkerColor = Colors.DarkGoldenrod
                            End If
                        Else
                            Dim span = doc.GetFullStatementSpan(lineNumber)
                            marker = New StatementMarker(span, lineNumber, Colors.Gold)
                            stMarkerProvider.AddStatementMarker(marker)
                        End If

                        If currentState = DebuggerState.Error Then
                            DiagramHelper.Helper.RunLater(m,
                                Sub()
                                    doc.EnsureLineVisible(line, True)
                                    Helper.MainWindow.ErrorPanel.ShowError(runner.Exception)
                                End Sub, 200)
                        Else
                            doc.EnsureLineVisible(line)
                        End If

                        doc.Activate()
                        lastDoc = doc
                        doc.ExecutionMarker = marker
                    End If
                End Sub)

        End Sub

        Public Sub ActivateIDE()
            Dim m = Helper.MainWindow
            m.Dispatcher.Invoke(
                  Sub()
                      If m.WindowState = System.Windows.WindowState.Minimized Then
                          m.WindowState = System.Windows.WindowState.Maximized
                      End If
                      m.Activate()
                      m.Focus()
                  End Sub)
        End Sub

        Public Sub StepInto()
            If _IsActive Then
                _ProgramEngine.StepInto()
            Else
                Run(breakOnStart:=True)
            End If
        End Sub

        Public Sub StepOut()
            If _IsActive Then
                _ProgramEngine.StepOut()
            Else
                Run(breakOnStart:=False)
            End If
        End Sub

        Friend Sub ShortStepOut()
            If _IsActive Then
                _ProgramEngine.ShortStepOut()
            Else
                Run(breakOnStart:=False)
            End If
        End Sub

        Public Sub [Continue]()
            _ProgramEngine.Continue()
        End Sub

        Public Sub StepOver()
            If _IsActive Then
                _ProgramEngine.StepOver()
            Else
                Run(breakOnStart:=True)
            End If
        End Sub

        Public Sub ShortStepOver()
            If _IsActive Then
                _ProgramEngine.ShortStepOver()
            Else
                Run(breakOnStart:=True)
            End If
        End Sub

        Public Function EvaluateExpression(ByRef text As String) As Library.Primitive?
            If Not IsActive OrElse text.Trim() = "" Then Return Nothing
            If isEvaluating Then Return Nothing

            isEvaluating = True
            Dim doc = Helper.MainWindow.ActiveDocument
            Dim runner = _ProgramEngine.GetEvaluationRunner(doc.File)
            If runner Is Nothing Then
                isEvaluating = False
                Return Nothing
            End If

            Dim span = doc.EditorControl.TextView.Selection.ActiveSnapshotSpan
            Dim line = span.Snapshot.GetLineFromPosition(span.Start)
            If line.Start < span.Start Then
                Dim c = span.Snapshot(span.Start)
                If c = "_" OrElse Char.IsLetter(c) Then
                    Dim prevChar = span.Snapshot(span.Start - 1)
                    If prevChar = "_"c OrElse prevChar = "."c OrElse Char.IsLetter(prevChar) Then
                        isEvaluating = False
                        Return Nothing
                    End If
                End If
            End If

            line = span.Snapshot.GetLineFromPosition(span.End)
            If line.End > span.End Then
                Dim c = span.Snapshot(span.End - 1)
                If c = "_" OrElse Char.IsLetter(c) Then
                    Dim nextChar = span.Snapshot(span.End)
                    If nextChar = "_"c OrElse nextChar = "."c OrElse Char.IsLetter(nextChar) Then
                        isEvaluating = False
                        Return Nothing
                    End If
                End If
            End If

            Dim result? As Library.Primitive
            Try
                result = runner.EvaluateExpression(text, doc.GetCurrentSubToken())
            Catch ex As Exception
                result = "Expression caused an error: " & ex.Message
            End Try

            isEvaluating = False
            Return result
        End Function

        Private Sub OnProgramTerminated()
            Dim m = Helper.MainWindow
            m.DisplayDebugCommands(False)
            m.Dispatcher.Invoke(
                Sub()
                    RemoveHandler _ProgramEngine.DebuggerStateChanged, AddressOf OnCurrentStateChanged
                    RemoveHandler Library.Program.ProgramTerminated, AddressOf OnProgramTerminated

                    _IsActive = False
                    _ProgramEngine.Dispose()
                    _ProgramEngine = Nothing
                    m.PopError.IsOpen = False
                    m.tabDesigner.IsEnabled = True

                    For Each view As Shell.MdiView In m.viewsControl.Items
                        Dim doc = view.Document
                        doc.EditorControl.Dispatcher.BeginInvoke(
                             Sub()
                                 doc.ReadOnly = False
                                 Dim stMarkerProvider = StatementMarkerProvider.GetStatementMarkerProvider(doc.EditorControl.TextView)
                                 Dim marker = doc.ExecutionMarker
                                 If marker IsNot Nothing Then
                                     doc.ExecutionMarker = Nothing
                                     If doc.Breakpoints.Contains(marker.LineNumber) Then
                                         marker.MarkerColor = Colors.Red
                                     Else
                                         stMarkerProvider.RemoveMarker(marker)
                                     End If
                                 End If
                             End Sub)
                    Next
                End Sub)
        End Sub

        Friend Shared Sub LockRunningDoc(doc As TextDocument)
            Dim docPath = doc.File.ToLower()
            For Each debugger In debuggers.Values
                If debugger.IsActive Then
                    For Each parser In debugger.ProgramEngine.Parsers
                        If parser.DocPath.ToLower() = docPath Then
                            doc.ReadOnly = True
                            Return
                        End If
                    Next
                End If
            Next
        End Sub
    End Class
End Namespace
