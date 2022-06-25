Imports System.Collections.Generic
Imports System.ComponentModel.Composition
Imports System.Windows.Input
Imports Microsoft.Nautilus.Core.Undo
Imports Microsoft.Nautilus.Text.Operations

Namespace Microsoft.Nautilus.Text.Editor
    <ComponentOptions(ComponentDiscoveryMode:=ComponentDiscoveryMode.Always)>
    Public NotInheritable Class DefaultKeyboardFilter
        Inherits KeyboardFilter

        Private _commandKeyBindingsMap As Dictionary(Of String, CommandKeyBinding)
        Private _commandHandlerMap As Dictionary(Of String, Action(Of IAvalonTextView))
        Private _commandKeyBindings As IEnumerable(Of CommandKeyBinding)
        Private _commandHandlers As IEnumerable(Of ImportInfo(Of Action(Of IAvalonTextView)))

        <Import>
        Public Property EditorOperationsProvider As IEditorOperationsProvider

        <Import>
        Public Property UndoHistoryRegistry As IUndoHistoryRegistry

        <Import("CommandKeyBinding")>
        Public Property CommandKeyBindings As IEnumerable(Of CommandKeyBinding)
            Get
                Return _commandKeyBindings
            End Get

            Set(value As IEnumerable(Of CommandKeyBinding))
                _commandKeyBindings = value
                _commandKeyBindingsMap = New Dictionary(Of String, CommandKeyBinding)
                For Each commandKeyBinding1 As CommandKeyBinding In _commandKeyBindings
                    Dim key As String = commandKeyBinding1.KeySequence.ToLowerInvariant()
                    If Not _commandKeyBindingsMap.ContainsKey(key) Then
                        _commandKeyBindingsMap(key) = commandKeyBinding1
                    End If
                Next
            End Set
        End Property

        <Import(GetType(CommandHandler))>
        Public Property CommandHandlers As IEnumerable(Of ImportInfo(Of Action(Of IAvalonTextView)))
            Get
                Return _commandHandlers
            End Get

            Set(value As IEnumerable(Of ImportInfo(Of Action(Of IAvalonTextView))))
                _commandHandlers = value
                _commandHandlerMap = New Dictionary(Of String, Action(Of IAvalonTextView))
                For Each commandHandler1 As ImportInfo(Of Action(Of IAvalonTextView)) In _commandHandlers
                    Dim key As String = TryCast(commandHandler1.Metadata("CommandName"), String)
                    If Not _commandHandlerMap.ContainsKey(key) Then
                        _commandHandlerMap(key) = commandHandler1.GetBoundValue()
                    End If
                Next
            End Set
        End Property

        Public Overrides Sub TextInput(textView As IAvalonTextView, args As TextCompositionEventArgs)
            If args.Text.Length > 0 Then
                Dim history As UndoHistory = UndoHistoryRegistry.GetHistory(textView.TextBuffer)
                Dim editorOperations As IEditorOperations = EditorOperationsProvider.GetEditorOperations(textView)
                editorOperations.InsertText(args.Text, history)
                textView.Caret.EnsureVisible()
            End If
            args.Handled = True
        End Sub

        Public Overrides Sub KeyDown(textView As IAvalonTextView, args As KeyEventArgs)

            Dim text1 As String = ""

            If (Keyboard.Modifiers And ModifierKeys.Control) = ModifierKeys.Control Then
                text1 &= "ctrl+"
            End If

            If (Keyboard.Modifiers And ModifierKeys.Alt) = ModifierKeys.Alt Then
                text1 &= "alt+"
            End If

            If (Keyboard.Modifiers And ModifierKeys.Shift) = ModifierKeys.Shift Then
                text1 &= "shift+"
            End If

            text1 &= args.Key.ToString().ToLowerInvariant()
            Dim keyBinding As CommandKeyBinding = Nothing

            If _commandKeyBindingsMap.TryGetValue(text1, keyBinding) Then
                Dim action As Action(Of IAvalonTextView) = Nothing
                If _commandHandlerMap.TryGetValue(keyBinding.Command, action) Then
                    action(textView)
                    args.Handled = True
                End If
            End If

        End Sub
    End Class
End Namespace
