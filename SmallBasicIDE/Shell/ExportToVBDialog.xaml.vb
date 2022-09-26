Imports Microsoft.SmallBasic
Imports Microsoft.SmallVisualBasic.Documents
Imports Microsoft.SmallVisualBasic.LanguageService
Imports System
Imports System.Diagnostics
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Windows
Imports System.Windows.Forms
Imports System.Windows.Markup
Imports System.Windows.Threading

Namespace Microsoft.SmallVisualBasic.Shell
    Public Partial Class ExportToVBDialog
        Inherits Window
        Implements IComponentConnector

        <Flags>
        Private Enum AssocF
            Init_NoRemapCLSID = &H1
            Init_ByExeName = &H2
            Open_ByExeName = &H2
            Init_DefaultToStar = &H4
            Init_DefaultToFolder = &H8
            NoUserSettings = &H10
            NoTruncate = &H20
            Verify = &H40
            RemapRunDll = &H80
            NoFixUps = &H100
            IgnoreBaseClass = &H200
        End Enum

        Private Enum AssocStr
            Command = 1
            Executable
            FriendlyDocName
            FriendlyAppName
            NoOpen
            ShellNewValue
            DDECommand
            DDEIfExec
            DDEApplication
            DDETopic
        End Enum

        Private document As TextDocument
        Private Shared previousLocation As String

        Public Sub New(document As TextDocument)
            If document Is Nothing Then
                Throw New ArgumentNullException("document")
            End If

            Me.document = document
            Me.InitializeComponent()
        End Sub

        Private Sub LaunchProject(projectPath As String)
            Dim pszOut As StringBuilder = New StringBuilder(1024)
            Dim pcchOut = 1024UI

            If AssocQueryString(AssocF.IgnoreBaseClass, AssocStr.Executable, ".vb", "open", pszOut, pcchOut) = 0 Then
                Me.continueButton.IsEnabled = False
                Me.cancelButton.IsEnabled = False
                Me.statusText.Text = ResourceHelper.GetString("Launching")
                Dim pro1 = Process.Start(projectPath)
                pro1.WaitForInputIdle()
                Dim tickCount = 0
                Dim dt As DispatcherTimer = New DispatcherTimer With {
                    .Interval = TimeSpan.FromMilliseconds(1000.0)
                }
                AddHandler dt.Tick, Sub()
                                        tickCount += 1

                                        If tickCount > 5 Then
                                            dt.Stop()
                                            Close()
                                        End If
                                    End Sub

                dt.Start()
            Else
                Me.statusPanel.Visibility = Visibility.Collapsed
                Me.notInstalledPanel.Visibility = Visibility.Visible
                Me.continueButton.IsEnabled = False
            End If
        End Sub

        Private Sub OnClickContinue(sender As Object, e As RoutedEventArgs)
            Me.statusPanel.Visibility = Visibility.Visible
            Me.statusText.Text = ResourceHelper.GetString("Converting")
            document.Errors.Clear()
            Dim compiler = CompilerService.Compile(document.Text, document.Errors)

            If document.Errors.Count = 0 Then
                Try
                    Dim visualBasicExporter As VisualBasicExporter = New VisualBasicExporter(compiler)
                    Dim projectName = If(Not document.IsNew, Path.GetFileNameWithoutExtension(document.FilePath), "Untitled")
                    Dim projectPath = visualBasicExporter.ExportToVisualBasicProject(projectName, Me.locationTextBox.Text)
                    LaunchProject(projectPath)
                Catch ex As Exception
                    System.Windows.MessageBox.Show(ex.Message)
                    Me.continueButton.IsEnabled = False
                End Try

                Return
            End If

            Me.statusPanel.Visibility = Visibility.Collapsed
            Me.errorPanel.Visibility = Visibility.Visible
            Me.continueButton.IsEnabled = False
        End Sub

        Private Sub OnClickBrowse(sender As Object, e As RoutedEventArgs)
            Dim folderBrowserDialog As FolderBrowserDialog = New FolderBrowserDialog()

            If Equals(previousLocation, Nothing) Then
                folderBrowserDialog.SelectedPath = Environment.GetFolderPath(Environment.SpecialFolder.Personal)
            Else
                folderBrowserDialog.SelectedPath = previousLocation
            End If

            If folderBrowserDialog.ShowDialog() = Forms.DialogResult.OK Then
                previousLocation = folderBrowserDialog.SelectedPath
                Me.locationTextBox.Text = folderBrowserDialog.SelectedPath
                Me.continueButton.IsEnabled = True
            End If
        End Sub

        Private Sub OnClickInstall(sender As Object, e As RoutedEventArgs)
            Process.Start("http://www.microsoft.com/express/vb/")
        End Sub

        <DllImport("Shlwapi.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
        Private Shared Function AssocQueryString(flags As AssocF, str As AssocStr, pszAssoc As String, pszExtra As String,
        <Out> pszOut As StringBuilder,
        <[In]>
        <Out> ByRef pcchOut As UInteger) As UInteger
        End Function
    End Class
End Namespace
