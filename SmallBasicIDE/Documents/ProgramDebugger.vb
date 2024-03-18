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
                Dim docPath = ProgramEngine.CurrentRunner.currentParser.DocPath
                Dim m = Helper.MainWindow
                Dim doc = If(docPath = "", m.ActiveDocument, m.OpenDocIfNot(docPath))
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
                AddHandler _ProgramEngine.DebuggerStateChanged, AddressOf OnDebuggerCurrentStateChanged
            End If
            Return _ProgramEngine
        End Function


        Public ReadOnly Property TextBuffer As ITextBuffer
            Get
                Return Document.TextBuffer
            End Get
        End Property

        Friend Function Evaluate(item As Completion.CompletionItem) As Library.Primitive?
            If item.ObjectName Is Nothing OrElse
                        item.ItemType = Completion.CompletionItemType.Control OrElse
                        item.ItemType = Completion.CompletionItemType.LocalVariable OrElse
                        item.ItemType = Completion.CompletionItemType.GlobalVariable Then
                Dim runner = _ProgramEngine.CurrentRunner
                Dim key = runner.GetKey(item.DefinitionIdintifier)
                If key = "" Then
                    key = item.DisplayName.ToLower()
                End If
                Return GetValue(key, runner.Fields)

            ElseIf item.ObjectName.ToLower() = "global" Then
                Dim key = item.DisplayName.ToLower()
                Return GetValue(key, _ProgramEngine.GlobalRunner.Fields)

            ElseIf item.ItemType = Completion.CompletionItemType.DynamicPropertyName Then
                Dim subName = Document.GetCurrentSubName()
                Dim arrKey = If(subName = "", item.ObjectName, $"{subName}.{item.ObjectName}").ToLower()
                Dim arr = GetValue(arrKey, _ProgramEngine.CurrentRunner.Fields)

                If arr.HasValue Then
                    Return Library.Primitive.GetArrayValue(arr.Value, item.DisplayName)
                ElseIf subName <> "" Then
                    arrKey = item.ObjectName.ToLower()
                    arr = GetValue(arrKey, _ProgramEngine.CurrentRunner.Fields)
                    If arr.HasValue Then
                        Return Library.Primitive.GetArrayValue(arr.Value, item.DisplayName)
                    End If
                End If

            Else
                Dim runner = _ProgramEngine.CurrentRunner
                Dim key = item.DisplayName.ToLower()
                Dim objectName As New Token() With {.Text = item.ObjectName, .Type = TokenType.Identifier}
                Dim type = runner.currentParser.SymbolTable.GetTypeInfo(objectName)

                Dim propInfo = TryCast(item.MemberInfo, Reflection.PropertyInfo)
                If propInfo IsNot Nothing Then
                    Return CType(propInfo.GetValue(Nothing, Nothing), Library.Primitive)
                Else
                    Dim name = GetValue(objectName.LCaseText, runner.Fields)
                    If name Is Nothing Then
                        Dim subName = Document.GetCurrentSubName()
                        If subName = "" Then Return Nothing
                        Dim objKey = $"{subName}.{item.ObjectName}".ToLower()
                        name = GetValue(objKey, runner.Fields)
                    End If

                    If name IsNot Nothing Then
                        Dim methodName = "get" & key
                        If Not type.Methods.ContainsKey(methodName) Then
                            type = runner.TypeInfoBag.Types("control")
                        End If
                        Dim methodInfo = type.Methods(methodName)
                        Return CType(methodInfo.Invoke(Nothing, New Object() {name}), Library.Primitive)
                    End If
                End If
            End If

            Return Nothing
        End Function

        Private Function GetValue(key As String, fields As Dictionary(Of String, Library.Primitive)) As Library.Primitive?
            If fields.ContainsKey(key) Then
                Return fields(key)
            Else
                Return Nothing
            End If
        End Function

        Public ReadOnly Property TextView As ITextView
            Get
                Return Document.EditorControl.TextView
            End Get
        End Property

        Public Sub Run(breakOnStart As Boolean)
            If IsActive Then
                _ProgramEngine.Continue()
            Else
                _ProgramEngine = CreateEngine()
                If _ProgramEngine IsNot Nothing Then
                    _ProgramEngine.RunProgram(breakOnStart)
                    Dim doc = Document
                    _ProgramEngine.CurrentRunner.Breakpoints = doc.Breakpoints
                    _IsActive = True
                    doc.ReadOnly = True
                End If
            End If
        End Sub

        Dim lastDoc As TextDocument

        Private Sub OnDebuggerCurrentStateChanged()
            Dim offset = _ProgramEngine.CurrentRunner.DocLineOffset
            Dim lineNumber = _ProgramEngine.CurrentLineNumber - offset
            Dim currentState = _ProgramEngine.CurrentDebuggerState

            Helper.MainWindow.Dispatcher.BeginInvoke(
                Sub()
                    Dim doc = If(lastDoc, Document)
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

                    If lastDoc IsNot Nothing Then
                        doc = Document
                        stMarkerProvider = doc.StatementMarkerProvider
                    End If

                    doc.ReadOnly = True

                    If currentState = DebuggerState.Paused Then
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

                        doc.ExecutionMarker = marker
                        doc.EnsureLineVisible(line)
                        doc.Activate()
                        lastDoc = doc
                    End If
                End Sub)

            ActivateDoc()
        End Sub

        Public Sub ActivateDoc()
            Dim m = Helper.MainWindow
            m.Dispatcher.BeginInvoke(
                  Sub()
                      If m.WindowState = System.Windows.WindowState.Minimized Then
                          m.WindowState = System.Windows.WindowState.Maximized
                      End If
                      m.Activate()
                      m.Focus()
                      m.tabCode.IsSelected = True
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
            If Not IsActive Then Return Nothing
            Dim subName = Helper.MainWindow.ActiveDocument.GetCurrentSubToken
            Return _ProgramEngine.CurrentRunner.EvaluateExpression(text, subName)
        End Function

        Private Sub OnProgramTerminated()
            RemoveHandler _ProgramEngine.DebuggerStateChanged, AddressOf OnDebuggerCurrentStateChanged
            RemoveHandler Library.Program.ProgramTerminated, AddressOf OnProgramTerminated

            _IsActive = False
            _ProgramEngine.Disopese()
            _ProgramEngine = Nothing

            For Each view As Shell.MdiView In Helper.MainWindow.viewsControl.Items
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
