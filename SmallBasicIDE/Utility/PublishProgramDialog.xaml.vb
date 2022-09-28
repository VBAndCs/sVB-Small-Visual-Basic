Imports Microsoft.SmallBasic
Imports Microsoft.SmallVisualBasic.com.smallbasic
Imports System
Imports System.Diagnostics
Imports System.Globalization
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input
Imports System.Windows.Markup

Namespace Microsoft.SmallVisualBasic.Utility
    Partial Public Class PublishProgramDialog
        Inherits Window
        Implements IComponentConnector

        Public Property HyperlinkRef As String
        Public Property ProgramId As String

        Public Sub New(programId As String)
            Me.InitializeComponent()
            Icon = Application.Current.MainWindow.Icon
            Me.ProgramId = programId
            HyperlinkRef = String.Format(CultureInfo.CurrentUICulture, ResourceHelper.GetString("LocationOfProgramHRef"), New Object(0) {Me.ProgramId})
            Me.hyperlink.Text = HyperlinkRef
            Me.saveId.Text = Me.ProgramId
        End Sub

        Private Sub OnClickClose(sender As Object, e As RoutedEventArgs)
            If Me.expander.IsExpanded Then
                Dim text As String = Me.titleTextBox.Text
                Dim text2 As String = Me.descriptionTextBox.Text
                Dim text3 As String = TryCast(CType(Me.categoryCombo.SelectedItem, TextBlock).Tag, String)

                If Not String.IsNullOrEmpty(text) OrElse Not String.IsNullOrEmpty(text2) OrElse Not String.IsNullOrEmpty(text3) Then
                    Cursor = Cursors.Wait

                    Try
                        Dim service As New Service()
                        service.PublishProgramDetails(ProgramId, If(text, ""), If(text2, ""), If(text3, "Miscellaneous"))
                    Catch ex As Exception
                        MessageBox.Show(ResourceHelper.GetString("FailedToUpdateProgramInfo"), ResourceHelper.GetString("Title"), ex.Message, NotificationButtons.Close, NotificationIcon.Error)
                    End Try
                End If
            End If

            Close()
        End Sub

        Private Sub OnClickHyperlink(sender As Object, e As RoutedEventArgs)
            Process.Start(HyperlinkRef)
        End Sub

        Private Sub OnExpanderChanged(sender As Object, e As RoutedEventArgs)
            If Me.expander.IsExpanded Then
                Me.titleTextBox.Focus()
                Me.closeButton.Content = ResourceHelper.GetString("UpdateAndClose")
            Else
                Me.closeButton.Content = ResourceHelper.GetString("Close")
            End If
        End Sub
    End Class
End Namespace
