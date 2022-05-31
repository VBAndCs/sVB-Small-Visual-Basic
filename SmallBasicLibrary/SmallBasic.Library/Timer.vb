Imports System.ComponentModel

Namespace Library
    ''' <summary>
    ''' The Timer object provides an easy way for doing something repeatedly with a constant interval between.
    ''' </summary>
    <SmallBasicType>
    Public NotInheritable Class Timer
        Public Const _maxInterval As Integer = 100000000
        Private Shared _interval As Integer
        Private Shared _threadTimer As Threading.Timer

        ''' <summary>
        ''' Gets or sets the interval (in milliseconds) specifying how often the timer should raise the Tick event.  This value can range from 10 to 100000000
        ''' </summary>
        Public Shared Property Interval As Primitive
            Get
                Return _interval
            End Get

            Set(Value As Primitive)
                _interval = Value
                _interval = Math.Max(10, Math.Min(_interval, 100000000))
                _threadTimer.Change(_interval, _interval)
            End Set
        End Property

        Private Shared Events As New EventHandlerList

        ''' <summary>
        ''' Raises an event when the timer ticks.
        ''' </summary>
        Public Shared Custom Event Tick As SmallBasicCallback
            AddHandler(Value As SmallBasicCallback)
                Dim Key = NameOf(Tick)
                Dim h = TryCast(Events(Key), SmallBasicCallback)
                If h IsNot Nothing Then Events.RemoveHandler(Key, h)
                Events.AddHandler(Key, Value)
            End AddHandler

            RemoveHandler(Value As SmallBasicCallback)
                Events.RemoveHandler(NameOf(Tick), Value)
            End RemoveHandler

            RaiseEvent()
                Dim h = TryCast(Events(NameOf(Tick)), SmallBasicCallback)
                If h IsNot Nothing Then h.Invoke()
            End RaiseEvent

        End Event

        Shared Sub New()
            _interval = 100000000
            _threadTimer = New Threading.Timer(AddressOf ThreadTimerCallback)
        End Sub

        ''' <summary>
        ''' Pauses the timer.  Tick events will not be raised.
        ''' </summary>
        Public Shared Sub Pause()
            _threadTimer.Change(-1, -1)
        End Sub

        ''' <summary>
        ''' Resumes the timer from a paused state.  Tick events will now be raised.
        ''' </summary>
        Public Shared Sub [Resume]()
            _threadTimer.Change(_interval, _interval)
        End Sub

        Private Shared Sub ThreadTimerCallback(tag As Object)
            RaiseEvent Tick()
        End Sub
    End Class
End Namespace
