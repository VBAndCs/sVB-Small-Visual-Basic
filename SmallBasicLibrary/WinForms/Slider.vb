Imports Microsoft.SmallVisualBasic.Library
Imports Wpf = System.Windows.Controls
Imports App = Microsoft.SmallVisualBasic.Library.Internal.SmallBasicApplication
Imports System.Windows

Namespace WinForms
    ''' <summary>
    ''' Represents the Slider control, that allows the user to choose a value within a range.
    ''' Use the Minimum and Maximun properties to set the slider range, use the Value property to set the current slide position, and use the OnSlide event to take action when the slider value changes.
    ''' You can use the form designer to add a slider to the form by dragging it from the toolbox.
    ''' It is also possible to use the Form.AddSlider method to create a new slider and add it to the form at runtime.
    ''' </summary>
    <SmallVisualBasicType>
    <HideFromIntellisense>
    Public NotInheritable Class Slider

        Private Shared Function GetSlider(sliderName As String) As Wpf.Slider
            Dim c = Control.GetControl(sliderName)
            Dim s = TryCast(c, Wpf.Slider)
            If s Is Nothing Then
                Throw New Exception($"{sliderName} is not a name of a Slider.")
            End If
            Return s
        End Function

        ''' <summary>
        ''' Gets or sets the value of current Slider
        ''' </summary>
        <ReturnValueType(VariableType.Double)>
        <ExProperty>
        Public Shared Function GetValue(sliderName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetValue = GetSlider(sliderName).Value
                    Catch ex As Exception
                        Control.ReportError(sliderName, "Value", ex)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetValue(sliderName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        GetSlider(sliderName).Value = value.AsDecimal()
                    Catch ex As Exception
                        Control.ReportPropertyError(sliderName, "Value", value, ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Gets or sets the minimum value of the current Slider. The default value is 0.
        ''' </summary>
        <ReturnValueType(VariableType.Double)>
        <ExProperty>
        Public Shared Function GetMinimum(sliderName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetMinimum = GetSlider(sliderName).Minimum
                    Catch ex As Exception
                        Control.ReportError(sliderName, "Minimum", ex)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetMinimum(sliderName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        GetSlider(sliderName).Minimum = value.AsDecimal()
                    Catch ex As Exception
                        Control.ReportPropertyError(sliderName, "Minimum", value, ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Gets or sets the maximum value of the current Slider. The default value is 100.
        ''' </summary>
        <ReturnValueType(VariableType.Double)>
        <ExProperty>
        Public Shared Function GetMaximum(sliderName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetMaximum = GetSlider(sliderName).Maximum
                    Catch ex As Exception
                        Control.ReportError(sliderName, "Maximum", ex)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetMaximum(sliderName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        GetSlider(sliderName).Maximum = value.AsDecimal()

                    Catch ex As Exception
                        Control.ReportPropertyError(sliderName, "Maximum", value, ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Gets or sets the distance between slide ticks.
        ''' Note that you can change the tics color by using the ForeColor property, which allows you to hid the ticks by setting the ForeColor property to Colors.None or Colors.Trabsparent.
        ''' </summary>
        <ReturnValueType(VariableType.Double)>
        <ExProperty>
        Public Shared Function GetTickFrequency(sliderName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetTickFrequency = GetSlider(sliderName).TickFrequency
                    Catch ex As Exception
                        Control.ReportError(sliderName, "TickFrequency", ex)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetTickFrequency(sliderName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        GetSlider(sliderName).TickFrequency = value.AsDecimal()
                    Catch ex As Exception
                        Control.ReportPropertyError(sliderName, "TickFrequency", value, ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Gets or sets whether or not the thumb movment snaps to tick marquees when the user slides it.
        ''' Set this property to True and set the TickFrequency property to a proper value to show marquees on the slider, and force the user to slide only to marquees positions.
        ''' The default value is False, which gives the user freedom to slide to positions between tha marquees.
        ''' </summary>
        <ReturnValueType(VariableType.Boolean)>
        <ExProperty>
        Public Shared Function GetSnapToTick(sliderName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetSnapToTick = GetSlider(sliderName).IsSnapToTickEnabled
                    Catch ex As Exception
                        Control.ReportError(sliderName, "SnapToTick", ex)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetSnapToTick(sliderName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        GetSlider(sliderName).IsSnapToTickEnabled = CBool(value)
                    Catch ex As Exception
                        Control.ReportPropertyError(sliderName, "SnapToTick", value, ex)
                    End Try
                End Sub)
        End Sub

        Public Shared ReadOnly TrackColorProperty As DependencyProperty =
                DependencyProperty.RegisterAttached("TrackColor",
                GetType(String), GetType(Slider),
                New PropertyMetadata("#FFE7EAEA"))

        ''' <summary>
        ''' Gets or sets the color used to draw the track of the slider bar.
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        <ExProperty>
        Public Shared Function GetTrackColor(sliderName As Primitive) As Primitive
            App.Invoke(
                  Sub()
                      Try
                          Dim c = GetSlider(sliderName)
                          GetTrackColor = CStr(c.GetValue(TrackColorProperty))
                      Catch ex As Exception
                          Control.ReportError(sliderName, "TrackColor", ex)
                      End Try
                  End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetTrackColor(sliderName As Primitive, value As Primitive)
            App.Invoke(
                   Sub()
                       Try
                           Dim c = GetSlider(sliderName)
                           c.SetValue(TrackColorProperty, value.AsString())
                           Dim border = TryCast(c.GetChild("TrackBackground"), Wpf.Border)
                           If border IsNot Nothing Then
                               border.Background = Color.GetBrush(value)
                           End If

                       Catch ex As Exception
                           Control.ReportPropertyError(sliderName, "TrackColor", value, ex)
                       End Try
                   End Sub)
        End Sub

        Public Shared ReadOnly ThumbColorProperty As DependencyProperty =
                DependencyProperty.RegisterAttached("ThumbColor",
                GetType(String), GetType(Slider),
                New PropertyMetadata("#FFD6D5D5"))

        ''' <summary>
        ''' Gets or sets the color used to draw the thunb of the slider.
        ''' </summary>
        <ReturnValueType(VariableType.Color)>
        <ExProperty>
        Public Shared Function GetThumbColor(sliderName As Primitive) As Primitive
            App.Invoke(
                  Sub()
                      Try
                          Dim c = GetSlider(sliderName)
                          GetThumbColor = CStr(c.GetValue(ThumbColorProperty))
                      Catch ex As Exception
                          Control.ReportError(sliderName, "ThumbColor", ex)
                      End Try
                  End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetThumbColor(sliderName As Primitive, value As Primitive)
            App.Invoke(
                   Sub()
                       Try
                           Dim c = GetSlider(sliderName)
                           c.SetValue(ThumbColorProperty, value.AsString())
                           Dim path = TryCast(c.GetChild("Background"), System.Windows.Shapes.Path)
                           If path IsNot Nothing Then
                               path.Fill = Color.GetBrush(value)
                           End If

                       Catch ex As Exception
                           Control.ReportPropertyError(sliderName, "ThumbColor", value, ex)
                       End Try
                   End Sub)
        End Sub


        ''' <summary>
        ''' Fired when the slider value changes.
        ''' </summary>
        Public Shared Custom Event OnSlide As SmallVisualBasicCallback
            AddHandler(handler As SmallVisualBasicCallback)
                Try
                    Dim _sender = GetSlider([Event].SenderControl)
                    AddHandler _sender.ValueChanged,
                        Sub(Sender As Wpf.Control, e As RoutedEventArgs)
                            [Event].EventsHandler(Sender, e, handler)
                        End Sub
                Catch ex As Exception
                    [Event].ShowErrorMessage(NameOf(OnSlide), ex)
                End Try
            End AddHandler

            RemoveHandler(handler As SmallVisualBasicCallback)
            End RemoveHandler

            RaiseEvent()
            End RaiseEvent
        End Event

    End Class
End Namespace