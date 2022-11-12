Imports Microsoft.SmallVisualBasic.Library
Imports Wpf = System.Windows.Controls
Imports App = Microsoft.SmallVisualBasic.Library.Internal.SmallBasicApplication
Imports System.Windows

Namespace WinForms
    ''' <summary>
    ''' Represents a DatePicker control, that allows the user to enter a date or pick it from the drop down calendar.
    ''' You can use the form designer to add a date picker to the form by dragging it from the toolbox.
    ''' It is also possible to use the Form.AddDatePicker method to create a new date picker and add it to the form at runtime.
    ''' </summary>
    <SmallVisualBasicType>
    <HideFromIntellisense>
    Public NotInheritable Class DatePicker

        Private Shared Function GetDatePicker(datePickerName As String) As Wpf.DatePicker
            Dim c = Control.GetControl(datePickerName)
            Dim dp = TryCast(c, Wpf.DatePicker)
            If dp Is Nothing Then
                Throw New Exception($"{datePickerName} is not a name of a DatePicker.")
            End If
            Return dp
        End Function

        ''' <summary>
        ''' Gets or sets the date that is selected by the DatePicker
        ''' </summary>
        <ReturnValueType(VariableType.Date)>
        <ExProperty>
        Public Shared Function GetSelectedDate(datePickerName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        Dim d = GetDatePicker(datePickerName).SelectedDate
                        GetSelectedDate = If(d.HasValue, New Primitive(d.Value.Ticks, NumberType.Date), "")
                    Catch ex As Exception
                        Control.ReportError(datePickerName, "SelectedDate", ex)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetSelectedDate(datePickerName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        GetDatePicker(datePickerName).SelectedDate = value.AsDate()
                    Catch ex As Exception
                        Control.ReportPropertyError(datePickerName, "SelectedDate", value, ex)
                    End Try
                End Sub)
        End Sub


        ''' <summary>
        ''' Fired when the text is changed.
        ''' </summary>
        Public Shared Custom Event OnSelection As SmallVisualBasicCallback
            AddHandler(handler As SmallVisualBasicCallback)
                Try
                    Dim _sender = GetDatePicker([Event].SenderControl)
                    AddHandler _sender.SelectedDateChanged, Sub(Sender As Wpf.Control, e As RoutedEventArgs) [Event].EventsHandler(CType(Sender, FrameworkElement), e, handler)
                Catch ex As Exception
                    [Event].ShowErrorMessage(NameOf(OnSelection), ex)
                End Try
            End AddHandler

            RemoveHandler(handler As SmallVisualBasicCallback)
            End RemoveHandler

            RaiseEvent()
            End RaiseEvent
        End Event

    End Class
End Namespace