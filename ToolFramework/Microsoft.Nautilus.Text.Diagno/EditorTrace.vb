
Namespace Microsoft.Nautilus.Text.Diagnostics
    Public NotInheritable Class EditorTrace

        Private Const ProviderName As String = "Nautilus Text Editor"
        Private Shared _traceProvider As New TraceProvider
        Private Shared ReadOnly _editorTraceProviderGuid As Guid
        Private Shared ReadOnly _keyDownEventGuid As Guid
        Private Shared _keyDownEvent As EtwEvent
        Private Shared ReadOnly _textRenderGuid As Guid
        Private Shared _textRenderEvent As EtwEvent
        Private Shared ReadOnly _textReplaceGuid As Guid
        Private Shared _textReplaceEvent As EtwEvent
        Private Shared ReadOnly _textInsertGuid As Guid
        Private Shared _textInsertEvent As EtwEvent
        Private Shared ReadOnly _textDeleteGuid As Guid
        Private Shared _textDeleteEvent As EtwEvent
        Private Shared ReadOnly _textCutGuid As Guid
        Private Shared _textCutEvent As EtwEvent
        Private Shared ReadOnly _textCopyGuid As Guid
        Private Shared _textCopyEvent As EtwEvent
        Private Shared ReadOnly _textPasteGuid As Guid
        Private Shared _textPasteEvent As EtwEvent
        Private Shared ReadOnly _textSelectGuid As Guid
        Private Shared _textSelectEvent As EtwEvent
        Private Shared ReadOnly _textPageDownGuid As Guid
        Private Shared _textPageDownEvent As EtwEvent
        Private Shared ReadOnly _textPageUpGuid As Guid
        Private Shared _textPageUpEvent As EtwEvent
        Private Shared ReadOnly _textScrollGuid As Guid
        Private Shared ReadOnly _textScrollEvent As EtwEvent
        Private Shared ReadOnly _textBufferChangeUndoGuid As Guid
        Private Shared ReadOnly _textBufferChangeUndoEvent As EtwEvent
        Private Shared ReadOnly _textBufferChangeRedoGuid As Guid
        Private Shared ReadOnly _textBufferChangeRedoEvent As EtwEvent
        Private Shared ReadOnly _beforeTextBufferChangeUndoGuid As Guid
        Private Shared ReadOnly _beforeTextBufferChangeUndoEvent As EtwEvent
        Private Shared ReadOnly _beforeTextBufferChangeRedoGuid As Guid
        Private Shared ReadOnly _beforeTextBufferChangeRedoEvent As EtwEvent

        Shared Sub New()
            _editorTraceProviderGuid = New Guid(3043426967UL, 33699, 16994, 139, 80, 63, 244, 199, 33, 207, 49)
            _keyDownEventGuid = New Guid(1002302878, 17450, 18149, 141, 138, 18, 104, 194, 27, 150, 89)
            _textRenderGuid = New Guid(3628515337UL, 6774, 19874, 150, 182, 22, 213, 211, 221, 101, 251)
            _textReplaceGuid = New Guid(833774530UL, 42169, 17807, 183, 16, 212, 195, 95, 187, 102, 222)
            _textInsertGuid = New Guid(3491398915UL, 16880, 18256, 147, 121, 144, 11, 29, 169, 167, 214)
            _textDeleteGuid = New Guid(4167499295UL, 29915, 19678, 189, 193, 162, 86, 253, 117, 81, 37)
            _textCutGuid = New Guid(3898605711UL, 33494, 20338, 135, 114, 206, 162, 57, 210, 40, 203)
            _textCopyGuid = New Guid(1835006663, 31321, 18947, 160, 154, 111, 55, 24, 205, 7, 227)
            _textPasteGuid = New Guid(3718857583UL, 40894, 18478, 155, 117, 22, 249, 213, 170, 4, 15)
            _textSelectGuid = New Guid(2105423529UL, 49953, 17717, 160, 3, 234, 124, 26, 35, 79, 90)
            _textPageDownGuid = New Guid(2856556316UL, 48150, 17355, 160, 225, 111, 47, 90, 75, 0, 14)
            _textPageUpGuid = New Guid(2467412871UL, 36543, 17632, 141, 207, 170, 11, 194, 19, 202, 188)
            _textScrollGuid = New Guid(864666809UL, 65398, 17528, 137, 69, 115, 130, 58, 220, 111, 41)
            _textBufferChangeUndoGuid = New Guid(2256936726UL, 28654, 16674, 142, 96, 105, 124, 134, 135, 112, 97)
            _textBufferChangeRedoGuid = New Guid(3936228402UL, 50, 16494, 175, 0, 10, Byte.MaxValue, 209, 155, 14, 11)
            _beforeTextBufferChangeUndoGuid = New Guid(4030765979UL, 43342, 18236, 165, 158, 162, 171, 158, 6, 188, 235)
            _beforeTextBufferChangeRedoGuid = New Guid(3988277835UL, 28224, 19870, 169, 108, 81, 90, 249, 224, 31, 163)
        End Sub

        Public Shared Sub TraceKeyDown()
            _traceProvider.RaiseEvent(_keyDownEvent, EventType.Checkpoint)
        End Sub

        Public Shared Sub TraceTextRenderStart()
            _traceProvider.RaiseEvent(_textRenderEvent, EventType.StartEvent)
        End Sub

        Public Shared Sub TraceTextRenderEnd()
            _traceProvider.RaiseEvent(_textRenderEvent, EventType.EndEvent)
        End Sub

        Public Shared Sub TraceTextInsertStart()
            _traceProvider.RaiseEvent(_textInsertEvent, EventType.StartEvent)
        End Sub

        Public Shared Sub TraceTextInsertEnd()
            _traceProvider.RaiseEvent(_textInsertEvent, EventType.EndEvent)
        End Sub

        Public Shared Sub TraceTextDeleteStart()
            _traceProvider.RaiseEvent(_textDeleteEvent, EventType.StartEvent)
        End Sub

        Public Shared Sub TraceTextDeleteEnd()
            _traceProvider.RaiseEvent(_textDeleteEvent, EventType.EndEvent)
        End Sub

        Public Shared Sub TraceTextReplaceStart()
            _traceProvider.RaiseEvent(_textReplaceEvent, EventType.StartEvent)
        End Sub

        Public Shared Sub TraceTextReplaceEnd()
            _traceProvider.RaiseEvent(_textReplaceEvent, EventType.EndEvent)
        End Sub

        Public Shared Sub TraceTextCutStart()
            _traceProvider.RaiseEvent(_textCutEvent, EventType.StartEvent)
        End Sub

        Public Shared Sub TraceTextCutEnd()
            _traceProvider.RaiseEvent(_textCutEvent, EventType.EndEvent)
        End Sub

        Public Shared Sub TraceTextCopyStart()
            _traceProvider.RaiseEvent(_textCopyEvent, EventType.StartEvent)
        End Sub

        Public Shared Sub TraceTextCopyEnd()
            _traceProvider.RaiseEvent(_textCopyEvent, EventType.EndEvent)
        End Sub

        Public Shared Sub TraceTextPasteStart()
            _traceProvider.RaiseEvent(_textPasteEvent, EventType.StartEvent)
        End Sub

        Public Shared Sub TraceTextPasteEnd()
            _traceProvider.RaiseEvent(_textPasteEvent, EventType.EndEvent)
        End Sub

        Public Shared Sub TraceTextSelectStart()
            _traceProvider.RaiseEvent(_textSelectEvent, EventType.StartEvent)
        End Sub

        Public Shared Sub TraceTextSelectEnd()
            _traceProvider.RaiseEvent(_textSelectEvent, EventType.EndEvent)
        End Sub

        Public Shared Sub TraceTextPageDownStart()
            _traceProvider.RaiseEvent(_textPageDownEvent, EventType.StartEvent)
        End Sub

        Public Shared Sub TraceTextPageDownEnd()
            _traceProvider.RaiseEvent(_textPageDownEvent, EventType.EndEvent)
        End Sub

        Public Shared Sub TraceTextPageUpStart()
            _traceProvider.RaiseEvent(_textPageUpEvent, EventType.StartEvent)
        End Sub

        Public Shared Sub TraceTextPageUpEnd()
            _traceProvider.RaiseEvent(_textPageUpEvent, EventType.EndEvent)
        End Sub

        Public Shared Sub TraceTextScrollStart()
            _traceProvider.RaiseEvent(_textScrollEvent, EventType.StartEvent)
        End Sub

        Public Shared Sub TraceTextScrollEnd()
            _traceProvider.RaiseEvent(_textScrollEvent, EventType.EndEvent)
        End Sub

        Public Shared Sub TraceTextBufferChangeUndoStart()
            _traceProvider.RaiseEvent(_textBufferChangeUndoEvent, EventType.StartEvent)
        End Sub

        Public Shared Sub TraceTextBufferChangeUndoEnd()
            _traceProvider.RaiseEvent(_textBufferChangeUndoEvent, EventType.EndEvent)
        End Sub

        Public Shared Sub TraceTextBufferChangeRedoStart()
            _traceProvider.RaiseEvent(_textBufferChangeRedoEvent, EventType.StartEvent)
        End Sub

        Public Shared Sub TraceTextBufferChangeRedoEnd()
            _traceProvider.RaiseEvent(_textBufferChangeRedoEvent, EventType.EndEvent)
        End Sub

        Public Shared Sub TraceBeforeTextBufferChangeUndoStart()
            _traceProvider.RaiseEvent(_beforeTextBufferChangeUndoEvent, EventType.StartEvent)
        End Sub

        Public Shared Sub TraceBeforeTextBufferChangeUndoEnd()
            _traceProvider.RaiseEvent(_beforeTextBufferChangeUndoEvent, EventType.EndEvent)
        End Sub

        Public Shared Sub TraceBeforeTextBufferChangeRedoStart()
            _traceProvider.RaiseEvent(_beforeTextBufferChangeRedoEvent, EventType.StartEvent)
        End Sub

        Public Shared Sub TraceBeforeTextBufferChangeRedoEnd()
            _traceProvider.RaiseEvent(_beforeTextBufferChangeRedoEvent, EventType.EndEvent)
        End Sub

        Public Shared Sub EnableTrace()
            _traceProvider.Enable()
        End Sub

        Public Shared Sub DisableTrace()
            _traceProvider.Disable()
        End Sub

    End Class
End Namespace
