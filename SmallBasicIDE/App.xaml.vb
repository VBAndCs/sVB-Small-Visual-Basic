Imports System
Imports System.Collections.Generic
Imports System.ComponentModel.Activation
Imports System.ComponentModel.Composition
Imports System.Globalization
Imports System.IO
Imports System.Threading
Imports System.Windows

Namespace Microsoft.SmallBasic
    Partial Public Class App
        Inherits Application

        Public Shared GlobalDomain As ComponentDomain

        Public Shared Property FlowDirection As FlowDirection

        Protected Overrides Sub OnStartup(e As StartupEventArgs)
            For Each arg In e.Args
                arg = arg.ToLowerInvariant().Trim()
                If File.Exists(arg) Then
                    Microsoft.SmallBasic.MainWindow.FilesToOpen.Add(arg)
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
            Dim text = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) & "\SmallBasic"

            If Not Directory.Exists(text) Then
                Directory.CreateDirectory(text)
            End If

            Dim catalogSourcesAggregator As CatalogSourcesAggregator = New CatalogSourcesAggregator()
            Dim directoryCatalogSource As DirectoryCatalogSource = New DirectoryCatalogSource()
            directoryCatalogSource.DirectoryPath = "."
            directoryCatalogSource.CacheFilePath = Path.Combine(text, "SB.exe.application.catalog")
            directoryCatalogSource.IsDynamic = False
            Dim item = directoryCatalogSource
            Dim directoryCatalogSource2 As DirectoryCatalogSource = New DirectoryCatalogSource()
            directoryCatalogSource2.DirectoryPath = "Settings"
            directoryCatalogSource2.CacheFilePath = Path.Combine(text, "SB.exe.settings.catalog")
            directoryCatalogSource2.IsDynamic = True
            Dim item2 = directoryCatalogSource2
            catalogSourcesAggregator.CatalogSources = New List(Of CatalogSource)()
            catalogSourcesAggregator.CatalogSources.Add(item)
            catalogSourcesAggregator.CatalogSources.Add(item2)
            Dim catalog As Catalog = catalogSourcesAggregator.CreateCatalog()
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
