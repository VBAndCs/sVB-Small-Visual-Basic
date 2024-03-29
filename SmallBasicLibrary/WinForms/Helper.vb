Imports Microsoft.SmallVisualBasic.Library
Imports App = Microsoft.SmallVisualBasic.Library.Internal.SmallBasicApplication

Module Helper

    Private exeFile As String = Process.GetCurrentProcess().MainModule.FileName
    Private exeDir As String = IO.Path.GetDirectoryName(exeFile)
    Private errorFile As String = IO.Path.Combine(exeDir, "RuntimeErrors.txt")
    Private ClearErrors As Boolean = True

    Public Sub ReportError(msg As String, ex As Exception)
        If app.IsDebugging Then
            Program.Exception = ex
            Return
        End If

        Try
            If ClearErrors Then
                ClearErrors = False
                File.DeleteFile(errorFile)
            End If
            File.AppendContents(errorFile, Now & vbCrLf & msg & vbCrLf)
            App.ShowErrorWindow(msg, ex.StackTrace)
        Catch
        End Try

    End Sub
End Module
