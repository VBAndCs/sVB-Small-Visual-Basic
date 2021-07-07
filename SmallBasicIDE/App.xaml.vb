Imports System
Imports System.Collections.Generic
Imports System.ComponentModel.Activation
Imports System.ComponentModel.Composition
Imports System.Globalization
Imports System.IO
Imports System.Threading
Imports System.Windows

Namespace Microsoft.SmallBasic
    Public Partial Class App
        Inherits Application

        Public Shared GlobalDomain As ComponentDomain
        Public Shared Property FlowDirection As FlowDirection

        Protected Overrides Sub OnStartup(e As StartupEventArgs)
            Dim args = e.Args

            For Each txt In args
                Dim text2 As String = txt.ToLowerInvariant().Trim()
                Dim num = text2.IndexOf(":")

                If num = -1 Then
                    Continue For
                End If

                Dim text3 = text2.Substring(0, num)
                Dim choice = text2.Substring(num + 1)

                Select Case text3
                    Case "/l", "/lang", "/language"
                        Dim cultureInfo = GetCultureInfo(choice)

                        If cultureInfo.IsNeutralCulture Then
                            Dim text4 = cultureInfo.Name

                            If Equals(text4, "zh-Hans") Then
                                text4 = "zh-CN"
                            ElseIf Equals(text4, "zh-Hant") Then
                                text4 = "zh-TW"
                            End If

                            cultureInfo = CultureInfo.CreateSpecificCulture(text4)
                        End If

                        Try
                            Thread.CurrentThread.CurrentCulture = cultureInfo
                        Catch
                        End Try

                        Thread.CurrentThread.CurrentUICulture = cultureInfo
                        Exit Select
                End Select
            Next

            Dim currentUICulture = CultureInfo.CurrentUICulture
            FlowDirection = If(currentUICulture.TextInfo.IsRightToLeft, FlowDirection.RightToLeft, FlowDirection.LeftToRight)
            CreateComponentDomain()
            'wndMain = New MainWindow
            GetMainWindow().Show()
            MyBase.OnStartup(e)
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


        'Friend Shared wndMain As MainWindow
    End Class
End Namespace
