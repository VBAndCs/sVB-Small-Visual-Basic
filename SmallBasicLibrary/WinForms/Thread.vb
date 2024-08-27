Imports Microsoft.SmallVisualBasic.Library

Namespace WinForms
    <SmallVisualBasicType>
    Public Class Thread

        ''' <summary>
        ''' Gets or sets the time in milliseconds that the main thread will be paused for, to allow the new thread to start and read the global variables.
        ''' The default value is 10, but you can increase it if the the thread needs to read many global variables.        
        ''' Use 0 if you don't need any delay.
        ''' </summary>
        <ReturnValueType(VariableType.Double)>
        Public Shared Property InitializationDelay As Primitive = New Primitive(10)

        ''' <summary>
        ''' This is not actually an event, but when you set its handler, it will be called immediately in a new thread to allow you to use execute a task in parallel to your normal code.
        ''' You can set this handler more than one time to create many threads, but take care not to create too many threads (at most 100 threads and the lesser the better) because that can make your system hang or block some of these threads.
        ''' Note also that you can't pass arguments to the handler, so you will need to use global variables.
        ''' Warning: Threads will not be created in debug mode to allow you to trace the code. In this case your program may be slower and some of its logic may not function correctly.
        ''' </summary>
        Public Shared Custom Event SubToRun As SmallVisualBasicCallback
            AddHandler(handler As SmallVisualBasicCallback)
                Dim t As New Threading.Thread(Sub() handler())
                t.IsBackground = True
                t.Start()
                If _InitializationDelay > 0 Then Program.Delay(_InitializationDelay)
            End AddHandler

            RemoveHandler(value As SmallVisualBasicCallback)
            End RemoveHandler
            RaiseEvent()
            End RaiseEvent
        End Event
    End Class
End Namespace