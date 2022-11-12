Imports App = Microsoft.SmallVisualBasic.Library.Internal.SmallBasicApplication
Imports Microsoft.SmallVisualBasic.Library
Imports System.Windows.Threading


Namespace WinForms
    ''' <summary>
    ''' Represents the WinTimer control, which provides an easy way for doing something repeatedly with a constant interval between.
    ''' Use the Interval property to set the time between ticks, and write the code you want to be executed in the OnTick event handler.
    ''' This object differs that the Timer object, as you can add many win timers as you want to each form in the project, while there is only a single Timer object that was suitable to SB code files.
    ''' You can't use the form designer to add Win timers to thee form, but you can use the Form.AddTimer method to add a new timer to the form.
    ''' </summary>
    <SmallVisualBasicType>
    <HideFromIntellisense>
    Public NotInheritable Class WinTimer
        Friend Shared Timers As New Dictionary(Of String, DispatcherTimer)

        Private Shared Function GetTimer(timerName As String) As DispatcherTimer
            If Timers.ContainsKey(timerName) Then
                Return Timers(timerName)
            Else
                Throw New Exception($"{timerName} is not a name of a Timer.")
            End If
        End Function


        ''' <summary>
        ''' Gets or sets the interval (in milliseconds) specifying how often the timer should raise the Tick event.  
        ''' Setting this property will start the timer and tick events will be raised.
        ''' </summary>
        <ReturnValueType(VariableType.Double)>
        <ExProperty>
        Public Shared Function GetInterval(timerName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetInterval = GetTimer(timerName).Interval.TotalMilliseconds
                    Catch ex As Exception
                        Control.ReportError(timerName, "Interval", ex)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetInterval(timerName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Dim t = GetTimer(timerName)
                        t.Interval = TimeSpan.FromMilliseconds(value)
                        t.Start()
                    Catch ex As Exception
                        Control.ReportError(timerName, "Interval", value, ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Pauses the timer. Tick events will not be raised.
        ''' </summary>
        <ExMethod>
        Public Shared Sub Pause(timerName As Primitive)
            App.Invoke(
                Sub()
                    Try
                        GetTimer(timerName).Stop()
                    Catch ex As Exception
                        Control.ReportSubError(timerName, "Pause", ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Resumes the timer from a paused state.  Tick events will now be raised.
        ''' </summary>
        <ExMethod>
        Public Shared Sub [Resume](timerName As Primitive)
            App.Invoke(
                Sub()
                    Try
                        GetTimer(timerName).Start()
                    Catch ex As Exception
                        Control.ReportSubError(timerName, "Pause", ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Fired when user releases the left mouse-button
        ''' </summary>
        Public Shared Custom Event OnTick As SmallVisualBasicCallback
            AddHandler(handler As SmallVisualBasicCallback)
                Try
                    Dim _sender = GetTimer([Event].SenderControl)
                    AddHandler _sender.Tick,
                        Sub(Sender As Object, e As EventArgs)
                            Try
                                [Event].SenderControl = CType(Sender, DispatcherTimer).Tag
                                Call handler()
                            Catch ex As Exception
                                ReportError($"The event handler sub `{handler.Method.Name}` caused this error: {ex.Message}", ex)
                            End Try
                        End Sub

                Catch ex As Exception
                    [Event].ShowErrorMessage(NameOf(OnTick), ex)
                End Try
            End AddHandler

            RemoveHandler(handler As SmallVisualBasicCallback)
            End RemoveHandler

            RaiseEvent()
            End RaiseEvent
        End Event

    End Class
End Namespace
