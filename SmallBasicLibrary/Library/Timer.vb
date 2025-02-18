Imports System.ComponentModel

Namespace Library
    ''' <summary>
    ''' The Timer object provides an easy way for doing something repeatedly with a constant interval between.
    ''' </summary>
    <SmallVisualBasicType>
    Public NotInheritable Class Timer
        Public Const _maxInterval As Integer = 100000000
        Private Shared _interval As Integer
        Private Shared _threadTimer As Threading.Timer

        ''' <summary>
        ''' Gets or sets the interval (in milliseconds) specifying how often the timer should raise the Tick event.  
        ''' This value can range from 1 to 100000000.
        ''' Note that setting the interval resumes the timer.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Property Interval As Primitive
            Get
                Return _interval
            End Get

            Set(Value As Primitive)
                _interval = Value
                _interval = System.Math.Max(1, System.Math.Min(_interval, 100000000))
                _threadTimer.Change(_interval, _interval)
            End Set
        End Property

        Private Shared tickHandler As SmallVisualBasicCallback

        ''' <summary>
        ''' Raises an event when the timer ticks.
        ''' </summary>
        Public Shared Custom Event Tick As SmallVisualBasicCallback
            AddHandler(Value As SmallVisualBasicCallback)
                tickHandler = Value
            End AddHandler

            RemoveHandler(Value As SmallVisualBasicCallback)
                tickHandler = Nothing
            End RemoveHandler

            RaiseEvent()
                If tickHandler IsNot Nothing Then tickHandler.Invoke()
            End RaiseEvent
        End Event

        ''' <summary>
        ''' Removes the Tick hamdler of the timer.
        ''' Ulternatively, you can set the Tick event to Nothing to remove its handler:
        ''' Timer.Tick = Nothing
        ''' </summary>
        Public Shared Sub RemoveTickHandler()
            tickHandler = Nothing
        End Sub


        Public Shared _isRunnning As Boolean

        ''' <summary>
        ''' Returns True is the timer is currently working (has an event handler and not paused), or False otherwise.
        ''' </summary>
        Public Shared ReadOnly Property IsRunnning As Boolean
            Get
                Return _isRunnning AndAlso tickHandler IsNot Nothing
            End Get
        End Property

        Shared Sub New()
            _interval = 100000000
            _threadTimer = New Threading.Timer(AddressOf ThreadTimerCallback)
            _isRunnning = True
        End Sub

        ''' <summary>
        ''' Pauses the timer.  Tick events will not be raised.
        ''' </summary>
        Public Shared Sub Pause()
            _threadTimer.Change(-1, -1)
            _isRunnning = False
        End Sub


        ''' <summary>
        ''' Resumes the timer from a paused state.  Tick events will now be raised.
        ''' </summary>
        Public Shared Sub [Resume]()
            If Not _isRunnning Then
                _threadTimer.Change(_interval, _interval)
                _isRunnning = True
            End If
        End Sub

        Private Shared Sub ThreadTimerCallback(tag As Object)
            If tickHandler IsNot Nothing Then tickHandler.Invoke()
        End Sub
    End Class
End Namespace
