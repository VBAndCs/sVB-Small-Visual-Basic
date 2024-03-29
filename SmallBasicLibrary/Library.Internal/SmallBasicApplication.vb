
Imports System.Reflection
Imports System.Threading
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Media
Imports System.Windows.Media.Imaging
Imports System.Windows.Threading

Namespace Library.Internal
    ''' <summary>
    ''' The Application class provides a Small Basic program with an application object.
    ''' </summary>
    Public NotInheritable Class SmallBasicApplication
        Private Shared _applicationThread As Thread
        Private Shared _application As Application
        Friend Shared mainThreadActions As Queue(Of InvokeHelper)
        Private Shared _pendingOperations As Integer

        ''' <summary>
        ''' Gets the dispatcher for the Small Basic application
        ''' </summary>
        Friend Shared ReadOnly Property Dispatcher As Dispatcher

        Friend Shared ReadOnly Property HasShutdown As Boolean
            Get
                Return Dispatcher.HasShutdownFinished
            End Get
        End Property

        Friend Shared IsDebugging As Boolean

        Shared Sub New()
            mainThreadActions = New Queue(Of InvokeHelper)
            _pendingOperations = 0
            Dim autoEvent As New AutoResetEvent(initialState:=False)
            _application = Application.Current

            If _application IsNot Nothing Then
                IsDebugging = True
                _application = Application.Current
                _Dispatcher = _application.Dispatcher
            Else
                IsDebugging = False
                _applicationThread = New Thread(
                     Sub()
                         Dim app As New Application()
                         autoEvent.Set()
                         _application = app
                         app.ShutdownMode = ShutdownMode.OnLastWindowClose
                         app.Run()
                     End Sub)

                _applicationThread.SetApartmentState(ApartmentState.STA)
                _applicationThread.Start()
                autoEvent.WaitOne()
                _Dispatcher = Dispatcher.FromThread(_applicationThread)
            End If

        End Sub

        Public Shared Sub BeginProgram()
            Program.IsTerminated = False
            AddHandler AppDomain.CurrentDomain.UnhandledException, AddressOf HandleException
        End Sub

        Friend Shared Function CropBitmap(original As BitmapSource, left As Primitive, top As Primitive, width As Primitive, height As Primitive) As BitmapSource
            If left < 0 Then left = 0
            If top < 0 Then top = 0
            If left + width > original.PixelWidth Then
                width = original.PixelWidth - left
            End If
            If top + height > original.PixelHeight Then
                height = original.PixelHeight - top
            End If

            Return CType(InvokeWithReturn(
                    Function() As CroppedBitmap
                        Return New CroppedBitmap(original, New Int32Rect(left, top, width, height))
                    End Function), BitmapSource)
        End Function


        Friend Shared Sub [End]()
            Invoke(
                Sub()
                    Timer.Pause()
                    RemoveHandler Timer.Tick, Nothing
                    Program.IsTerminated = True

                    If IsDebugging Then
                        If GraphicsWindow._window IsNot Nothing Then
                            Try
                                GraphicsWindow._windowCreated = False
                                GraphicsWindow._window.Close()
                            Catch
                            End Try
                        End If

                        Stack._stackMap = New Dictionary(Of Primitive, Stack(Of Primitive))
                        TextWindow.Hide()
                        WinForms.Forms.ForceCloseAll()
                        Return
                    End If

                    If GraphicsWindow._window Is Nothing Then
                        Try
                            GraphicsWindow._window.Close()
                        Catch
                        End Try
                    End If

                    _application.Shutdown()
                    _Dispatcher.InvokeShutdown()
                    Process.GetCurrentProcess().Kill()
                End Sub)
        End Sub

        Friend Shared Sub BeginInvoke(invokeDelegate As InvokeHelper)
            _pendingOperations += 1
            _Dispatcher.BeginInvoke(DispatcherPriority.Render, invokeDelegate)
            If _pendingOperations > 40 Then
                ClearDispatcherQueue()
                _pendingOperations = 0
            End If
        End Sub

        Friend Shared Function GetEntryAssemblyPath() As String
            Return Assembly.GetEntryAssembly().Location
        End Function

        Friend Shared Function GetResourceUri(resourceName As String) As Uri
            Return New Uri($"pack://application:,,/SmallVisualBasicLibrary;component/Resources/{resourceName}")
        End Function

        Private Shared Sub HandleException(sender As Object, e As UnhandledExceptionEventArgs)
            Dim ex = TryCast(e.ExceptionObject, Exception)
            If ex IsNot Nothing Then
                ShowErrorWindow(ex.Message, ex.StackTrace)
            End If
        End Sub

        Friend Shared Sub ShowErrorWindow(message As String, stackTrace As String)
            Invoke(
                Sub()
                    Dim wind As New Window With {
                       .Title = "Error in Small Visual Basic Program",
                       .SizeToContent = SizeToContent.WidthAndHeight,
                       .WindowStyle = WindowStyle.SingleBorderWindow,
                       .ShowInTaskbar = False,
                       .ResizeMode = ResizeMode.NoResize,
                       .Topmost = True,
                       .WindowStartupLocation = WindowStartupLocation.CenterScreen,
                       .Content =
                             Function() As StackPanel
                                 Dim stackPanel As New StackPanel With {.Margin = New Thickness(4.0)}
                                 stackPanel.Children.Add(
                                       Function() As DockPanel
                                           Dim dockPanel1 As New DockPanel
                                           dockPanel1.Children.Add(New Image With {
                                                      .Source = New BitmapImage(New Uri("pack://application:,,/SmallVisualBasicLibrary;component/Resources/Error.png")),
                                                      .Width = 48.0,
                                                      .Height = 48.0,
                                                      .Margin = New Thickness(8.0),
                                                      .VerticalAlignment = VerticalAlignment.Top
                                           })
                                           dockPanel1.Children.Add(
                                                   Function() As StackPanel
                                                       Dim stackPanel1 As New StackPanel
                                                       stackPanel1.Children.Add(New TextBlock With {
                                                                .Text = message & vbCrLf &
                                                                            "You can press Ctrl+F5 in Small Visual Basic IDE to debug the project where ot will stop at the line that causes the error." & vbCrLf &
                                                                            "Press Esc to close this window and end the program.",
                                                                .TextWrapping = TextWrapping.Wrap,
                                                                .Margin = New Thickness(4.0),
                                                                .MaxWidth = 650,
                                                                .FontSize = 16.0,
                                                                .Foreground = New SolidColorBrush(Color.FromArgb(Byte.MaxValue, 0, 42, 181))
                                                       })
                                                       Return stackPanel1
                                                   End Function())
                                           Return dockPanel1
                                       End Function())
                                 stackPanel.Children.Add(New TextBox With {
                                       .Text = stackTrace,
                                       .FontFamily = New FontFamily("Consolas"),
                                       .IsReadOnly = True,
                                       .MaxWidth = 650,
                                       .MaxHeight = 500,
                                       .HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                                       .VerticalScrollBarVisibility = ScrollBarVisibility.Auto
                                 })
                                 Return stackPanel
                             End Function()
                    }
                    AddHandler wind.PreviewKeyDown,
                        Sub(obj As Object, info As Input.KeyEventArgs) If info.Key = Input.Key.Escape Then wind.Close()
                    wind.ShowDialog()
                    Program.End()
                End Sub)
        End Sub

        Friend Shared Sub Invoke(invokeDelegate As InvokeHelper)
            _Dispatcher.Invoke(DispatcherPriority.Render, invokeDelegate)
        End Sub

        Friend Shared Function InvokeWithReturn(invokeDelegate As InvokeHelperWithReturn) As Object
            Return _Dispatcher.Invoke(DispatcherPriority.Render, invokeDelegate)
        End Function

        Friend Shared Sub ClearDispatcherQueue()
            Try
                _Dispatcher.Invoke(DispatcherPriority.Background, CType(Function(param0 As Object) Nothing, DispatcherOperationCallback), Nothing)
            Catch

            End Try
        End Sub

        Public Shared Sub EndProgram()
            If Not GraphicsWindow._windowVisible AndAlso Timer.Interval = 100000000 Then
                Invoke(Sub()
                           _application.Shutdown()
                           _Dispatcher.InvokeShutdown()
                       End Sub)

                If AppDomain.CurrentDomain.FriendlyName <> "Debuggee" Then
                    Process.GetCurrentProcess().Kill()
                End If
            End If
        End Sub

        Protected Overrides Sub Finalize()
            MyBase.Finalize()
        End Sub
    End Class
End Namespace
