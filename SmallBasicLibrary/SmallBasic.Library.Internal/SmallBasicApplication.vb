
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

        Shared Sub New()
            mainThreadActions = New Queue(Of InvokeHelper)
            _pendingOperations = 0
            Dim autoEvent As New AutoResetEvent(initialState:=False)
            _applicationThread = New Thread(
                Sub()
                    Dim app As New Application
                    autoEvent.[Set]()
                    _application = app
                    app.ShutdownMode = ShutdownMode.OnLastWindowClose
                    app.Run()
                End Sub)

            _applicationThread.SetApartmentState(ApartmentState.STA)
            _applicationThread.Start()
            autoEvent.WaitOne()
            _dispatcher = Dispatcher.FromThread(_applicationThread)
        End Sub

        Public Shared Sub BeginProgram()
            AddHandler AppDomain.CurrentDomain.UnhandledException, AddressOf HandleException
        End Sub

        Friend Shared Function CropBitmap(original As BitmapSource, left As Primitive, top As Primitive, width As Primitive, height As Primitive) As BitmapSource
            If CBool((left < 0)) Then
                left = 0
            End If
            If CBool((top < 0)) Then
                top = 0
            End If
            If CBool((left + width > original.PixelWidth)) Then
                width = original.PixelWidth - left
            End If
            If CBool((top + height > original.PixelHeight)) Then
                height = original.PixelHeight - top
            End If

            Return CType(InvokeWithReturn(Function() As System.Windows.Media.Imaging.CroppedBitmap
                                              Return New CroppedBitmap(original, New Int32Rect(left, top, width, height))
                                          End Function), BitmapSource)
        End Function

        Friend Shared Sub [End]()
            Invoke(Sub()
                       _application.Shutdown()
                       _Dispatcher.InvokeShutdown()
                   End Sub)
            If AppDomain.CurrentDomain.FriendlyName <> "Debuggee" Then
                Process.GetCurrentProcess().Kill()
            End If
        End Sub

        Friend Shared Sub BeginInvoke(invokeDelegate As InvokeHelper)
            _pendingOperations += 1
            _dispatcher.BeginInvoke(DispatcherPriority.Render, invokeDelegate)
            If _pendingOperations > 40 Then
                ClearDispatcherQueue()
                _pendingOperations = 0
            End If
        End Sub

        Friend Shared Function GetEntryAssemblyPath() As String
            Return Assembly.GetEntryAssembly().Location
        End Function

        Friend Shared Function GetResourceUri(resourceName As String) As Uri
            Return New Uri($"pack://application:,,/SmallBasicLibrary;component/Resources/{resourceName}")
        End Function

        Private Shared Sub HandleException(sender As Object, e As UnhandledExceptionEventArgs)
            Dim exception1 As Exception = TryCast(e.ExceptionObject, Exception)
            If exception1 IsNot Nothing Then
                Invoke(
                    Sub()
                        Dim window1 As New Window With {
                           .Title = "Error in Small Basic Program",
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
                                                          .Source = New BitmapImage(New Uri("pack://application:,,/SmallBasicLibrary;component/Resources/Error.png")),
                                                          .Width = 48.0,
                                                          .Height = 48.0,
                                                          .Margin = New Thickness(8.0),
                                                          .VerticalAlignment = VerticalAlignment.Top
                                               })
                                               dockPanel1.Children.Add(
                                                       Function() As StackPanel
                                                           Dim stackPanel1 As New StackPanel
                                                           stackPanel1.Children.Add(New TextBlock With {
                                                                    .Text = exception1.Message,
                                                                    .TextWrapping = TextWrapping.Wrap,
                                                                    .Margin = New Thickness(4.0),
                                                                    .FontSize = 16.0,
                                                                    .Foreground = New SolidColorBrush(Color.FromArgb(Byte.MaxValue, 0, 42, 181))
                                                           })
                                                           stackPanel1.Children.Add(New TextBox With {
                                                                    .Text = exception1.StackTrace,
                                                                    .FontFamily = New FontFamily("Consolas"),
                                                                    .IsReadOnly = True,
                                                                    .MaxWidth = 650,
                                                                    .Margin = New Thickness(4.0),
                                                                    .HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                                                                    .VerticalScrollBarVisibility = ScrollBarVisibility.Auto
                                                           })
                                                           Return stackPanel1
                                                       End Function())
                                               Return dockPanel1
                                           End Function())
                                     stackPanel.Children.Add(
                                          Function() As StackPanel
                                              Dim stackPanel2 As New StackPanel With {
                                                    .Orientation = Orientation.Horizontal,
                                                    .HorizontalAlignment = HorizontalAlignment.Right
                                              }
                                              stackPanel2.Children.Add(New Button With {
                                                     .Content = "OK",
                                                     .Margin = New Thickness(4.0),
                                                      .Width = 60.0,
                                                      .IsCancel = True
                                               })
                                              Return stackPanel2
                                          End Function())
                                     Return stackPanel
                                 End Function()
                        }
                        window1.ShowDialog()
                    End Sub)
            End If

        End Sub

        Friend Shared Sub Invoke(invokeDelegate As InvokeHelper)
            _dispatcher.Invoke(DispatcherPriority.Render, invokeDelegate)
        End Sub

        Friend Shared Function InvokeWithReturn(invokeDelegate As InvokeHelperWithReturn) As Object
            Return _dispatcher.Invoke(DispatcherPriority.Render, invokeDelegate)
        End Function

        Friend Shared Sub ClearDispatcherQueue()
            Try
                _dispatcher.Invoke(DispatcherPriority.Background, CType(Function(param0 As Object) Nothing, DispatcherOperationCallback), Nothing)
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
    End Class
End Namespace
