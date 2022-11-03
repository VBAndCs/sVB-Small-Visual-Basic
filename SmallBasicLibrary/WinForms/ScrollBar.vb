Imports Microsoft.SmallVisualBasic.Library
Imports Wpf = System.Windows.Controls
Imports App = Microsoft.SmallVisualBasic.Library.Internal.SmallBasicApplication
Imports System.Windows

Namespace WinForms
    <SmallVisualBasicType>
    <HideFromIntellisense>
    Public NotInheritable Class ScrollBar

        Private Shared Function GetScrollBar(scrollBarName As String) As Wpf.Primitives.ScrollBar
            Dim c = Control.GetControl(scrollBarName)
            Dim s = TryCast(c, Wpf.Primitives.ScrollBar)
            If s Is Nothing Then
                Throw New Exception($"{scrollBarName} is not a name of a ScrollBar.")
            End If
            Return s
        End Function

        ''' <summary>
        ''' Gets or sets the value of current ScrollBar
        ''' </summary>
        <ReturnValueType(VariableType.Double)>
        <ExProperty>
        Public Shared Function GetValue(scrollBarName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetValue = GetScrollBar(scrollBarName).Value
                    Catch ex As Exception
                        Control.ReportError(scrollBarName, "Value", ex)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetValue(scrollBarName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        GetScrollBar(scrollBarName).Value = value.AsDecimal()
                    Catch ex As Exception
                        Control.ReportPropertyError(scrollBarName, "Value", value, ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Gets or sets the minimum value of the current ScrollBar. The default value is 0.
        ''' </summary>
        <ReturnValueType(VariableType.Double)>
        <ExProperty>
        Public Shared Function GetMinimum(scrollBarName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetMinimum = GetScrollBar(scrollBarName).Minimum
                    Catch ex As Exception
                        Control.ReportError(scrollBarName, "Minimum", ex)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetMinimum(scrollBarName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Dim s = GetScrollBar(scrollBarName)
                        s.Minimum = value.AsDecimal()
                        s.SmallChange = (s.Maximum - s.Minimum) / 20
                        s.LargeChange = 2 * s.SmallChange
                    Catch ex As Exception
                        Control.ReportPropertyError(scrollBarName, "Minimum", value, ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Gets or sets the maximum value of the current ScrollBar. The default value is 100.
        ''' </summary>
        <ReturnValueType(VariableType.Double)>
        <ExProperty>
        Public Shared Function GetMaximum(scrollBarName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetMaximum = GetScrollBar(scrollBarName).Maximum
                    Catch ex As Exception
                        Control.ReportError(scrollBarName, "Maximum", ex)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetMaximum(scrollBarName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Dim s = GetScrollBar(scrollBarName)
                        s.Maximum = value.AsDecimal()
                        s.SmallChange = (s.Maximum - s.Minimum) / 20
                        s.LargeChange = 2 * s.SmallChange
                    Catch ex As Exception
                        Control.ReportPropertyError(scrollBarName, "Maximum", value, ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Fired when the scrollBar value changes.
        ''' </summary>
        Public Shared Custom Event OnScroll As SmallBasicCallback
            AddHandler(handler As SmallBasicCallback)
                Try
                    Dim _sender = GetScrollBar([Event].SenderControl)
                    AddHandler _sender.ValueChanged,
                        Sub(Sender As Wpf.Control, e As RoutedEventArgs)
                            [Event].EventsHandler(Sender, e, handler)
                        End Sub

                Catch ex As Exception
                    [Event].ShowErrorMessage(NameOf(OnScroll), ex)
                End Try
            End AddHandler

            RemoveHandler(handler As SmallBasicCallback)
            End RemoveHandler

            RaiseEvent()
            End RaiseEvent
        End Event

    End Class
End Namespace