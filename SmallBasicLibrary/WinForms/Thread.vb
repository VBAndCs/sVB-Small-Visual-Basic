Imports Microsoft.SmallVisualBasic.Library

Namespace WinForms
    <SmallVisualBasicType>
    Public Class Thread
        ''' <summary>
        ''' This is not actually an event, but when you set its handler, it will be called immediately in a new thread to allow you to use execute a task in parallel to your normal code.
        ''' You can set this handler more than one time to create many threads, but take care not to create too many threads (at most 100 threads and the lesser the better) because that can make your system hang or block some of these threads.
        ''' Note also that you can't pass arguments to the handler, so you will need to use global variables. This method will pause your code for 10ms to allow the new thread to start and read the global variables, but if this is not enough, you may need to call Program.Delay to increase the delay.
        ''' </summary>
        Public Shared Custom Event SubToRun As SmallVisualBasicCallback
            AddHandler(handler As SmallVisualBasicCallback)
                Dim t As New Threading.Thread(Sub() handler())
                t.Start()
                Program.Delay(10)
            End AddHandler

            RemoveHandler(value As SmallVisualBasicCallback)
            End RemoveHandler
            RaiseEvent()
            End RaiseEvent
        End Event
    End Class
End Namespace