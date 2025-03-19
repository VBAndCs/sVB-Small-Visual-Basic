﻿
Imports Microsoft.SmallVisualBasic.Documents
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Media.Animation

Namespace Microsoft.SmallVisualBasic.Shell
    Public Class MdiView
        Inherits ContentControl

        Public ReadOnly Property MdiViews As MdiViewsControl

        Public Overrides Sub OnApplyTemplate()
            MyBase.OnApplyTemplate()
            _CmbControlNames = Me.Template.FindName("CmbControlNames", Me)
            _CmbEventNames = Me.Template.FindName("CmbEventNames", Me)
            _NumZoom = Me.Template.FindName("NumZoom", Me)
            _NumZoom.Value = TextDocument.ScaleFactor * 100
            _MdiViews = GetParent(Of MdiViewsControl)(_CmbEventNames)
        End Sub

        Public Shared ReadOnly IsSelectedProperty As DependencyProperty = DependencyProperty.Register("IsSelected", GetType(Boolean), GetType(MdiView))
        Public Shared ReadOnly CaretPositionTextProperty As DependencyProperty = DependencyProperty.Register("CaretPositionText", GetType(String), GetType(MdiView))
        Private oldWidth As Double = Double.NaN
        Private oldHeight As Double = Double.NaN

        Public ReadOnly Property CmbControlNames As ComboBox

        Public ReadOnly Property CmbEventNames As ComboBox

        Public Property NumZoom As WpfDialogs.DoubleUpDown

        Friend Property FreezeCmbEvents As Boolean


        Public Property CaretPositionText As String
            Get
                Return CStr(GetValue(CaretPositionTextProperty))
            End Get

            Set(value As String)
                SetValue(CaretPositionTextProperty, value)
            End Set
        End Property

        Dim _document As TextDocument
        Friend Const FormSuffix As String = " (Form)"

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

            Set(value As Boolean)
                SetValue(IsSelectedProperty, value)
            End Set
        End Property

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

        Private Sub DoubleAnimateProperty([property] As DependencyProperty, start As Double, [end] As Double, timespan As Double)
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

        Friend Sub SelectEventName(form As String, controlName As String, eventName As String, freezeCmbEvents As Boolean)
            _FreezeCmbEvents = freezeCmbEvents
            _CmbControlNames.SelectedItem = If(form = controlName, form & FormSuffix, controlName)
            _CmbEventNames.SelectedItem = eventName
            MdiViews.SelectHandlers(Me, controlName, eventName)
            _FreezeCmbEvents = False
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
