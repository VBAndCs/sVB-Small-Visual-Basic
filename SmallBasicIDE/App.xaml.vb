Imports System
Imports System.Collections.Generic
Imports System.ComponentModel.Activation
Imports System.ComponentModel.Composition
Imports System.Globalization
Imports System.IO
Imports System.Threading
Imports System.Windows

Namespace Microsoft.SmallVisualBasic
    Partial Public Class App
        Inherits Application

        Public Shared GlobalDomain As ComponentDomain

        Public Shared Property FlowDirection As FlowDirection

        Protected Overrides Sub OnStartup(e As StartupEventArgs)
            For Each arg In e.Args
                arg = arg.ToLowerInvariant().Trim()
                If File.Exists(arg) Then
                    Microsoft.SmallVisualBasic.MainWindow.FilesToOpen.Add(arg)
                    Continue For
                End If

                Dim num = arg.IndexOf(":")
                If num = -1 Then Continue For

                Dim command = arg.Substring(0, num)
                Dim choice = arg.Substring(num + 1)

                Select Case command
                    Case "/l", "/lang", "/language"
                        Dim cultureInfo = GetCultureInfo(choice)

                        If cultureInfo.IsNeutralCulture Then
                            Dim cultureName = cultureInfo.Name

                            If Equals(cultureName, "zh-Hans") Then
                                cultureName = "zh-CN"
                            ElseIf Equals(cultureName, "zh-Hant") Then
                                cultureName = "zh-TW"
                            End If

                            cultureInfo = CultureInfo.CreateSpecificCulture(cultureName)
                        End If

                        Try
                            Thread.CurrentThread.CurrentCulture = cultureInfo
                        Catch
                        End Try

                        Thread.CurrentThread.CurrentUICulture = cultureInfo
                End Select
            Next

            Dim currentUICulture = CultureInfo.CurrentUICulture
            FlowDirection = If(currentUICulture.TextInfo.IsRightToLeft, FlowDirection.RightToLeft, FlowDirection.LeftToRight)
            CreateComponentDomain()
            AddHandler AppDomain.CurrentDomain.UnhandledException, AddressOf CatchAppExceptions

            Me.MainWindow = GetMainWindow()
            Me.MainWindow.Show()
            MyBase.OnStartup(e)
        End Sub

        Private Sub CatchAppExceptions(sender As Object, e As UnhandledExceptionEventArgs)
            Dim x = TryCast(e.ExceptionObject, Exception)
            Console.WriteLine(If(x IsNot Nothing, x.Message, e.ExceptionObject.ToString()))
        End Sub

        Private Function GetMainWindow() As Window
            Return GlobalDomain.GetBoundValue(Of Window)("MainWindow")
        End Function

        Private Sub CreateComponentDomain()
            LoaderXamlServices.EnsureActivatorsInitialized()
            Dim sbDir = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) & "\SmallVisualBasic"

            If Not Directory.Exists(sbDir) Then
                Directory.CreateDirectory(sbDir)
            End If

            Dim catalogSourcesAggregator As New CatalogSourcesAggregator()
            catalogSourcesAggregator.CatalogSources = New List(Of CatalogSource) From {
                    New DirectoryCatalogSource With {
                    .DirectoryPath = ".",
                    .CacheFilePath = Path.Combine(sbDir, "sVB.exe.application.catalog"),
                    .IsDynamic = False
                },
                New DirectoryCatalogSource With {
                    .DirectoryPath = "Settings",
                    .CacheFilePath = Path.Combine(sbDir, "sVB.exe.settings.catalog"),
                    .IsDynamic = True
                }
            }
            Dim catalog = catalogSourcesAggregator.CreateCatalog()
            GlobalDomain = New ComponentDomain(New CatalogResolver(catalog))
        End Sub

        Private Function GetCultureInfo(choice As String) As CultureInfo
            Try
                Return CultureInfo.GetCultureInfo(choice)
            Catch
                Return CultureInfo.CurrentCulture
            End Try
        End Function

        Public Shared CntxMenuIsOpened As Boolean

        Sub CntxMenu_Opened()
            CntxMenuIsOpened = True
        End Sub

        Sub CntxMenu_Closed()
            CntxMenuIsOpened = False
        End Sub

    End Class
End Namespace
