Imports System
Imports System.ComponentModel
Imports System.Globalization
Imports System.Reflection
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input
Imports System.Windows.Interop
Imports System.Windows.Markup
Imports System.Windows.Media
Imports System.Windows.Media.Imaging

Namespace Microsoft.SmallVisualBasic.Utility
    Partial Public Class MessageBox
        Inherits Window
        Implements INotifyPropertyChanged, IComponentConnector

        Friend selectedButton As NotificationButtons
        Private notificationButtonsField As NotificationButtons
        Private defaultButtonField As NotificationButtons
        Private notificationIconField As NotificationIcon
        Private isCancelableField As Boolean = True
        Private descriptionField As String
        Private optionalContentField As Object
        Private footerField As Object
        Private closeInvokedInternally As Boolean
        Private hwndSource As HwndSource

        Public Property Description As String
            Get
                Return descriptionField
            End Get
            Set(value As String)
                descriptionField = value
                RaisePropertyChangeEvent("Description")
            End Set
        End Property

        Public Property OptionalContent As Object
            Get
                Return optionalContentField
            End Get
            Set(value As Object)
                optionalContentField = value
                RaisePropertyChangeEvent("OptionalContent")
            End Set
        End Property

        Public Property Footer As Object
            Get
                Return footerField
            End Get
            Set(value As Object)
                footerField = value
                RaisePropertyChangeEvent("Footer")
            End Set
        End Property

        Public Property IsCancelable As Boolean
            Get
                Return isCancelableField
            End Get
            Set(value As Boolean)
                isCancelableField = value
                RaisePropertyChangeEvent("IsCancelable")
            End Set
        End Property

        Public Property NotificationButtons As NotificationButtons
            Get
                Return notificationButtonsField
            End Get
            Set(value As NotificationButtons)
                notificationButtonsField = value
                RaisePropertyChangeEvent("NotificationButtons")
            End Set
        End Property

        Public Property DefaultButton As NotificationButtons
            Get
                Return defaultButtonField
            End Get
            Set(value As NotificationButtons)
                defaultButtonField = value
                RaisePropertyChangeEvent("DefaultButton")
            End Set
        End Property

        Public Property NotificationIcon As NotificationIcon
            Get
                Return notificationIconField
            End Get
            Set(value As NotificationIcon)
                notificationIconField = value
                RaisePropertyChangeEvent("NotificationIcon")
            End Set
        End Property

        Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

        Friend Sub New()
            notificationButtonsField = NotificationButtons.OK
            defaultButtonField = NotificationButtons.OK
            Me.InitializeComponent()
        End Sub

        Public Function Display() As NotificationButtons
            DisplayHelper()
            Return selectedButton
        End Function

        Public Shared Function Show(description As String, title As String, optionalContent As Object, buttons As NotificationButtons, icon As NotificationIcon) As NotificationButtons
            Dim messageBox As MessageBox = New MessageBox()
            messageBox.Description = description
            messageBox.Title = title
            messageBox.OptionalContent = optionalContent
            messageBox.NotificationButtons = buttons
            messageBox.NotificationIcon = icon
            Return messageBox.Display()
        End Function

        Private Sub DisplayHelper()
            Owner = Application.Current.MainWindow
            selectedButton = NotificationButtons.Cancel
            Icon = Application.Current.MainWindow.Icon
            ComposeUI()
            ShowDialog()
        End Sub

        Friend Sub ComposeUI()
            SetWindowCloseButtonState()

            If optionalContentField Is Nothing Then
                Me.optionalContentControl.Visibility = Visibility.Collapsed
            Else
                Me.optionalContentControl.Visibility = Visibility.Visible
            End If

            If footerField Is Nothing Then
                Me.footerPanel.Visibility = Visibility.Collapsed
            Else
                Me.footerPanel.Visibility = Visibility.Visible
            End If

            Dim iconImage As ImageSource = GetIconImage()

            If iconImage IsNot Nothing Then
                Me.iconImageControl.Source = iconImage
                Me.iconImageControl.Visibility = Visibility.Visible
            ElseIf NotificationIcon = NotificationIcon.Custom Then
                Me.iconImageControl.Visibility = Visibility.Visible
            Else
                Me.iconImageControl.Visibility = Visibility.Collapsed
            End If

            Dim button = Me.closeButton
            Dim button2 = Me.cancelButton
            Dim button3 = Me.retryButton
            Dim button4 = Me.noButton
            Dim button5 = Me.yesButton

            Me.okButton.Visibility = Visibility.Collapsed
            button5.Visibility = Visibility.Collapsed
            button4.Visibility = Visibility.Collapsed
            button3.Visibility = Visibility.Collapsed
            button2.Visibility = Visibility.Collapsed
            button.Visibility = Visibility.Collapsed

            If (notificationButtonsField And NotificationButtons.Close) = NotificationButtons.Close Then
                Me.closeButton.Visibility = Visibility.Visible
            End If

            If (notificationButtonsField And NotificationButtons.Cancel) = NotificationButtons.Cancel Then
                Me.cancelButton.Visibility = Visibility.Visible
            End If

            If (notificationButtonsField And NotificationButtons.Retry) = NotificationButtons.Retry Then
                Me.retryButton.Visibility = Visibility.Visible
            End If

            If (notificationButtonsField And NotificationButtons.No) = NotificationButtons.No Then
                Me.noButton.Visibility = Visibility.Visible
            End If

            If (notificationButtonsField And NotificationButtons.Yes) = NotificationButtons.Yes Then
                Me.yesButton.Visibility = Visibility.Visible
            End If

            If (notificationButtonsField And NotificationButtons.OK) = NotificationButtons.OK Then
                Me.okButton.Visibility = Visibility.Visible
            End If

            Dim button6 As Button = Me.okButton

            If defaultButtonField = NotificationButtons.Close Then
                button6 = Me.closeButton
            ElseIf defaultButtonField = NotificationButtons.Cancel Then
                button6 = Me.cancelButton
            ElseIf defaultButtonField = NotificationButtons.Retry Then
                button6 = Me.retryButton
            ElseIf defaultButtonField = NotificationButtons.No Then
                button6 = Me.noButton
            ElseIf defaultButtonField = NotificationButtons.Yes Then
                button6 = Me.yesButton
            ElseIf defaultButtonField = NotificationButtons.OK Then
                button6 = Me.okButton
            End If

            button6.IsDefault = True
            button6.Focus()
        End Sub

        Private Sub SetWindowCloseButtonState()
            If hwndSource IsNot Nothing Then
                If isCancelableField Then
                    Dim systemMenu = GetSystemMenu(hwndSource.Handle, 0)
                    EnableMenuItem(systemMenu, 61536, 0)
                Else
                    Dim systemMenu2 = GetSystemMenu(hwndSource.Handle, 0)
                    EnableMenuItem(systemMenu2, 61536, 1)
                End If
            End If
        End Sub

        Private Sub RaisePropertyChangeEvent(propertyName As String)
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propertyName))
        End Sub

        Private Function GetIconImage() As ImageSource
            Dim text As String = Nothing
            Dim result As BitmapImage = Nothing

            Select Case notificationIconField
                Case NotificationIcon.Information
                    text = "Information.png"
                Case NotificationIcon.Warning
                    text = "Warning.png"
                Case NotificationIcon.Error
                    text = "Error.png"
                Case NotificationIcon.Shield
                    text = "Shield.png"
            End Select

            If Not Equals(text, Nothing) Then
                Dim name As String = Assembly.GetExecutingAssembly().GetName().Name
                Dim uriSource As Uri = New Uri(String.Format(CultureInfo.CurrentUICulture, "/{0};component/Resources/{1}", New Object(1) {name, text}), UriKind.Relative)
                result = New BitmapImage(uriSource)
            End If

            Return result
        End Function

        Private Sub OnButtonClick(sender As Object, e As RoutedEventArgs)
            If sender Is Me.closeButton Then
                selectedButton = NotificationButtons.Close
            ElseIf sender Is Me.cancelButton Then
                selectedButton = NotificationButtons.Cancel
            ElseIf sender Is Me.retryButton Then
                selectedButton = NotificationButtons.Retry
            ElseIf sender Is Me.noButton Then
                selectedButton = NotificationButtons.No
            ElseIf sender Is Me.yesButton Then
                selectedButton = NotificationButtons.Yes
            ElseIf sender Is Me.okButton Then
                selectedButton = NotificationButtons.OK
            End If

            closeInvokedInternally = True
            Close()
            closeInvokedInternally = False
        End Sub

        Private Sub OnCloseButtonClick(sender As Object, e As RoutedEventArgs)
            closeInvokedInternally = True
            Close()
            closeInvokedInternally = False
        End Sub

        Protected Overrides Sub OnClosing(e As CancelEventArgs)
            If Not closeInvokedInternally AndAlso Not isCancelableField Then
                e.Cancel = True
            End If

            MyBase.OnClosing(e)
        End Sub

        Protected Overrides Sub OnClosed(e As EventArgs)
            If hwndSource IsNot Nothing Then
                hwndSource.Dispose()
                hwndSource = Nothing
            End If

            MyBase.OnClosed(e)
        End Sub

        Protected Overrides Sub OnKeyDown(e As KeyEventArgs)
            If e.Key = Key.Cancel OrElse e.Key = Key.Escape Then
                Close()
            End If

            MyBase.OnKeyDown(e)
        End Sub

        Protected Overrides Sub OnSourceInitialized(e As EventArgs)
            hwndSource = CType(PresentationSource.FromVisual(Me), HwndSource)
            SetWindowCloseButtonState()
            MyBase.OnSourceInitialized(e)
        End Sub

    End Class
End Namespace
