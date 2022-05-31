Imports Microsoft.Nautilus.Text.Editor
Imports Microsoft.SmallBasic.Documents
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Media.Animation

Namespace Microsoft.SmallBasic.Shell
    Public Class MdiView
        Inherits ContentControl

        Public Shared ReadOnly IsSelectedProperty As DependencyProperty = DependencyProperty.Register("IsSelected", GetType(Boolean), GetType(MdiView))
        Public Shared ReadOnly CaretPositionTextProperty As DependencyProperty = DependencyProperty.Register("CaretPositionText", GetType(String), GetType(MdiView))
        Private oldWidth As Double = Double.NaN
        Private oldHeight As Double = Double.NaN

        Public ReadOnly Property CmbControlNames As ComboBox
        Public ReadOnly Property CmbEventNames As ComboBox
        Friend Property FreezeCmbEvents As Boolean

        Public Overrides Sub OnApplyTemplate()
            MyBase.OnApplyTemplate()
            _CmbControlNames = Me.Template.FindName("CmbControlNames", Me)
            _CmbEventNames = Me.Template.FindName("CmbEventNames", Me)
        End Sub

        Public Property CaretPositionText As String
            Get
                Return CStr(GetValue(CaretPositionTextProperty))
            End Get

            Set(value As String)
                SetValue(CaretPositionTextProperty, value)
            End Set
        End Property

        Dim _document As TextDocument

        Public Property Document As TextDocument
            Get
                Return _document
            End Get

            Set(value As TextDocument)
                _document = value
                _document.MdiView = Me
            End Set
        End Property


        Public Property IsSelected As Boolean
            Get
                Return CBool(GetValue(IsSelectedProperty))
            End Get

            Set(ByVal value As Boolean)
                SetValue(IsSelectedProperty, value)
            End Set
        End Property

        Protected Overrides Function MeasureOverride(ByVal constraint As Size) As Size
            If Double.IsNaN(Width) Then
                Return New Size(640.0, 480.0)
            End If

            Return MyBase.MeasureOverride(constraint)
        End Function

        Public Sub ResetOldWidthAndHeight()
            oldWidth = Double.NaN
            oldHeight = Double.NaN
        End Sub

        Private Sub Collapse()
            oldWidth = ActualWidth
            oldHeight = ActualHeight
            DoubleAnimateProperty(WidthProperty, oldWidth, MinWidth, 200.0)
            DoubleAnimateProperty(HeightProperty, oldHeight, MinHeight, 200.0)
        End Sub

        Private Sub Expand()
            If Not Double.IsNaN(oldWidth) AndAlso Not Double.IsNaN(oldHeight) Then
                DoubleAnimateProperty(WidthProperty, ActualWidth, oldWidth, 200.0)
                DoubleAnimateProperty(HeightProperty, ActualHeight, oldHeight, 200.0)
            End If
        End Sub

        Private Sub DoubleAnimateProperty(ByVal [property] As DependencyProperty, ByVal start As Double, ByVal [end] As Double, ByVal timespan As Double)
            Dim doubleAnimation As DoubleAnimation = New DoubleAnimation(start, [end], New Duration(System.TimeSpan.FromMilliseconds(timespan)))
            doubleAnimation.DecelerationRatio = 0.9
            Dim doubleAnimation2 = doubleAnimation
            AddHandler doubleAnimation2.Completed,
                Sub()
                    BeginAnimation([property], Nothing)
                    SetValue([property], [end])
                End Sub

            BeginAnimation([property], doubleAnimation2)
        End Sub

        ''' <summary>
        ''' Select an Item in the EventNames ConboBox without executing the SelectionChanged Event
        ''' </summary>
        ''' <param name="name"></param>
        Friend Sub SelectEventName(name As String)
            _FreezeCmbEvents = True
            If name = "-1" Then
                _CmbEventNames.SelectedIndex = -1
            Else
                _CmbEventNames.SelectedItem = name
            End If
            _FreezeCmbEvents = False
        End Sub
    End Class
End Namespace
