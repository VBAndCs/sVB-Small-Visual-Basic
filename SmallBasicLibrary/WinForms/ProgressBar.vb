﻿Imports Microsoft.SmallVisualBasic.Library
Imports Wpf = System.Windows.Controls
Imports App = Microsoft.SmallVisualBasic.Library.Internal.SmallBasicApplication

Namespace WinForms
    ''' <summary>
    ''' Represents the ProgressBar control, that indicates how much of the task has done.
    ''' Use the Minimum and Maximun properties to set the progress range, and use the Value property to set the current progress.
    ''' You can use the form designer to add a progress bar to the form by dragging it from the toolbox.
    ''' It is also possible to use the Form.AddProgressBar method to create a new progress bar and add it to the form at runtime.
    ''' </summary>
    <SmallVisualBasicType>
    <HideFromIntellisense>
    Public NotInheritable Class ProgressBar

        Private Shared Function GetProgressBar(progressBarName As String) As Wpf.ProgressBar
            Dim c = Control.GetControl(progressBarName)
            Dim pb = TryCast(c, Wpf.ProgressBar)
            If pb Is Nothing Then
                Throw New Exception($"{progressBarName} is not a name of a ProgressBar.")
            End If
            Return pb
        End Function

        ''' <summary>
        ''' Gets or sets the progress value in current ProgressBar
        ''' </summary>
        <ReturnValueType(VariableType.Double)>
        <ExProperty>
        Public Shared Function GetValue(progressBarName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetValue = GetProgressBar(progressBarName).Value
                    Catch ex As Exception
                        Control.ReportError(progressBarName, "Value", ex)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetValue(progressBarName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Dim pb = GetProgressBar(progressBarName)
                        pb.Dispatcher.Invoke(
                            Sub() pb.Value = value.AsDecimal(),
                            System.Windows.Threading.DispatcherPriority.Render)

                    Catch ex As Exception
                        Control.ReportPropertyError(progressBarName, "Value", value, ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Gets or sets the progress minimum value in current ProgressBar. The default value is 0.
        ''' </summary>
        <ReturnValueType(VariableType.Double)>
        <ExProperty>
        Public Shared Function GetMinimum(progressBarName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        GetMinimum = GetProgressBar(progressBarName).Minimum
                    Catch ex As Exception
                        Control.ReportError(progressBarName, "Minimum", ex)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetMinimum(progressBarName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        GetProgressBar(progressBarName).Minimum = value.AsDecimal()
                    Catch ex As Exception
                        Control.ReportPropertyError(progressBarName, "Minimum", value, ex)
                    End Try
                End Sub)
        End Sub

        ''' <summary>
        ''' Gets or sets the progress maximum value in current ProgressBar. The default value is 100.
        ''' If you set the max value to 0, this will mean you don't know when the progress will end (indeterminate), so, it will show an infinitly moving marquee.
        ''' </summary>
        <ReturnValueType(VariableType.Double)>
        <ExProperty>
        Public Shared Function GetMaximum(progressBarName As Primitive) As Primitive
            App.Invoke(
                Sub()
                    Try
                        Dim pb = GetProgressBar(progressBarName)
                        GetMaximum = If(pb.IsIndeterminate, 0, pb.Maximum)
                    Catch ex As Exception
                        Control.ReportError(progressBarName, "Maximum", ex)
                    End Try
                End Sub)
        End Function

        <ExProperty>
        Public Shared Sub SetMaximum(progressBarName As Primitive, value As Primitive)
            App.Invoke(
                Sub()
                    Try
                        Dim pb = GetProgressBar(progressBarName)
                        Dim dbl = value.AsDecimal()

                        If dbl <= 0 Then
                            pb.IsIndeterminate = True
                        Else
                            pb.IsIndeterminate = False
                            pb.Maximum = dbl
                        End If

                    Catch ex As Exception
                        Control.ReportPropertyError(progressBarName, "Maximum", value, ex)
                    End Try
                End Sub)
        End Sub

    End Class
End Namespace