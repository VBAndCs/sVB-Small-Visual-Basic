Imports System.Collections.ObjectModel
Imports System.IO

Public Class ProjectExplorer
    Inherits Explorer

    Dim _projDir As String
    Dim projFiles As New ObservableCollection(Of ProjFileInfo)

    Public Event FileNameChanged(oldFileName As String, newFileName As String)

    Public Property ProjectDirectory As String
        Get
            Return _projDir
        End Get

        Set()
            Dim fileName = ""
            If Value = "" OrElse Not (Directory.Exists(Value) OrElse File.Exists(Value)) Then
                _projDir = Value
                Title = "New Project Files"
                projFiles.Clear()
                Return
            End If

            If File.GetAttributes(Value) = FileAttributes.Directory Then
                Value = Value.ToLower()
            Else
                fileName = Value
                Value = Path.GetDirectoryName(Value).ToLower()
            End If

            If Value = _projDir Then
                If fileName <> "" Then SelectItem(fileName)
                Return
            End If

            _projDir = Value
            Title = $"{GetProjectName(_projDir)} Project Files"
            Dim files = Directory.GetFiles(_projDir, "*.xaml")
            FreezListFiles = True
            projFiles.Clear()

            If files.Count > 0 Then
                For Each f In files
                    projFiles.Add(New ProjFileInfo(f.ToLower()))
                Next
            End If

            Dim globFile = Path.Combine(_projDir, "global.sb")
            If File.Exists(globFile) Then projFiles.Add(New ProjFileInfo(globFile))
            FreezListFiles = False

            If fileName = "" Then
                SelectedIndex = 0
            Else
                SelectItem(fileName)
            End If
        End Set
    End Property

    Private Sub SelectItem(fileName As String)
        For i = 0 To projFiles.Count - 1
            If projFiles(i).FilePath = fileName Then
                FilesList.SelectedIndex = i
                Return
            End If
        Next
    End Sub

    Private Function GetProjectName(projDir As String) As Object
        ' The job of this function is to restore the normal case project name from the lower case one.
        If projDir = "" Then Return "New"
        Dim dir = Path.GetDirectoryName(projDir)
        Dim dirs = Directory.GetDirectories(dir, Path.GetFileName(projDir))
        Return "`" & Path.GetFileName(dirs(0)) & "`"
    End Function

    Protected Overrides ReadOnly Property ItemsSource As Specialized.INotifyCollectionChanged
        Get
            Return projFiles
        End Get
    End Property

    Protected Overrides Sub OnSelectionChanged()
        Designer.SwitchTo(CType(FilesList.SelectedItem, ProjFileInfo).FilePath)
    End Sub

    Protected Overrides Sub OnDeleteItem()
        If FilesList.SelectedIndex = -1 Then
            Beep()
            Return
        End If

        If MsgBox("This action will delete this form and its code files from the project folder and move them to the recycle bin. Are you sure you want to do that?", vbYesNo Or MsgBoxStyle.Exclamation, "Warning") = MsgBoxResult.No Then
            Return
        End If

        Try
            Dim i = FilesList.SelectedIndex
            Dim projFileInfo = CType(FilesList.SelectedItem, ProjFileInfo)
            Dim fileName = projFileInfo.FilePath
            FileIO.FileSystem.DeleteFile(fileName, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin)

            fileName = fileName.Substring(0, fileName.Length - 4) & "sb"
            FileIO.FileSystem.DeleteFile(fileName, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin)

            fileName &= ".gen"
            FileIO.FileSystem.DeleteFile(fileName, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin)

            projFiles.Remove(projFileInfo)

            Dim n = FilesList.Items.Count
            If n = 0 Then Return
            FilesList.SelectedIndex = If(i >= n, n - 1, i)

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Protected Overrides Function OnCommit(newName As String) As Boolean
        Dim newFile = Path.Combine(_projDir, newName & ".xaml")
        Dim oldFile = CType(FilesList.SelectedItem, ProjFileInfo).FilePath

        If newName.ToLower() = Path.GetFileNameWithoutExtension(oldFile).ToLower() Then Return True

        If File.Exists(newFile) Then
            MsgBox($"The project folder already contains a file named `{newName}`", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical)
            Return False
        End If

        If File.Exists(oldFile) Then
            File.Move(oldFile, newFile)

            Dim progFileInfo = CType(FilesList.SelectedItem, ProjFileInfo)
            Dim i = projFiles.IndexOf(progFileInfo)
            projFiles.RemoveAt(i)
            progFileInfo.FilePath = newFile
            projFiles.Insert(i, progFileInfo)

            oldFile = oldFile.Substring(0, oldFile.Length - 4) & "sb"
            newFile = newFile.Substring(0, newFile.Length - 4) & "sb"
            If File.Exists(oldFile) And Not File.Exists(newFile) Then
                File.Move(oldFile, newFile)
            End If

            RaiseEvent FileNameChanged(oldFile, newFile)

            oldFile &= ".gen"
            newFile &= ".gen"
            If File.Exists(oldFile) AndAlso Not File.Exists(newFile) Then
                File.Move(oldFile, newFile)
            End If

            Return True
        Else
                MsgBox($"The file `{oldFile}` doesn't exist in project folder", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical)
            Return False
        End If
    End Function

End Class

Public Class ProjFileInfo
    Public Sub New(filePath As String)
        Me.FilePath = filePath
    End Sub

    Public Property FilePath As String

    Public Overrides Function ToString() As String
        Return " ● " & Path.GetFileName(_FilePath)
    End Function
End Class