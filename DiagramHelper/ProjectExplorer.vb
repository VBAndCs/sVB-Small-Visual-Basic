Imports System.Collections.ObjectModel
Imports System.IO

Public Class ProjectExplorer
    Inherits Explorer

    Dim _projDir As String
    Dim projFiles As New ObservableCollection(Of ProjFileInfo)
    Public Shared CurrentProject As String

    Public Event FileNameChanged(oldFileName As String, newFileName As String)
    Public Event FileDeleted()

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

            If File.GetAttributes(Value) <> FileAttributes.Directory Then
                fileName = Value
                Value = Path.GetDirectoryName(Value)
            End If

            If Value.ToLower() = _projDir.ToLower() Then
                If fileName <> "" Then
                    If Path.GetFileNameWithoutExtension(fileName).ToLower() = "global" Then
                        SelectItem(Path.Combine(_projDir, "Global.sb"))
                    Else
                        SelectItem(fileName)
                    End If
                End If
                Return
            End If

            _projDir = Value
            CurrentProject = _projDir
            Title = $"{GetProjectName(_projDir)} Project Files"
            FreezListFiles = True
            projFiles.Clear()

            Dim files = Directory.GetFiles(_projDir, "*.xaml")
            If files.Count > 0 Then
                For Each f In files
                    Dim fName = Helper.GetFormNameFromXaml(f)
                    If fName = "" Then Continue For
                    projFiles.Add(New ProjFileInfo(f))
                Next
            End If

            Dim globFile = Path.Combine(_projDir, "Global.sb")
            projFiles.Add(New ProjFileInfo(globFile))

            FreezListFiles = False

            If fileName = "" Then
                SelectedIndex = 0
            Else
                SelectItem(fileName)
            End If

            Designer.RecentDirectory = _projDir
        End Set
    End Property

    Private Sub SelectItem(fileName As String)
        Dim f = fileName.ToLower()
        If SelectedFile.ToLower() = f Then Return

        For i = 0 To projFiles.Count - 1
            If projFiles(i).FilePath.ToLower() = f Then
                FilesList.SelectedIndex = i
                Return
            End If
        Next

        If File.Exists(fileName) Then
            projFiles.Add(New ProjFileInfo(fileName))
            FilesList.SelectedIndex = projFiles.Count - 1
        Else
            FilesList.SelectedIndex = -1
        End If
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

    Public ReadOnly Property SelectedFile As String
        Get
            If FilesList.SelectedIndex = -1 Then Return ""
            Return CType(FilesList.SelectedItem, ProjFileInfo).FilePath
        End Get
    End Property

    Protected Overrides Sub OnSelectionChanged()
        Dim filePath = SelectedFile
        If Path.GetFileNameWithoutExtension(filePath).ToLower() = "global" Then
            Designer.SwitchTo(Helper.GlobalFileName)
            Designer.CurrentPage.CodeFile = filePath
        Else
            Designer.SwitchTo(filePath)
        End If
    End Sub

    Protected Overrides Sub OnDeleteItem()
        If FilesList.SelectedIndex = -1 Then
            Beep()
            Return
        End If

        If MsgBox(
                    "This action will delete this form and its code files from the project folder and move them to the recycle bin. Are you sure you want to do that?",
                     vbYesNo Or MsgBoxStyle.Exclamation Or MsgBoxStyle.DefaultButton2,
                     "Warning"
                 ) = MsgBoxResult.No Then
            Return
        End If

        Dim i = FilesList.SelectedIndex
        Dim projFileInfo = CType(FilesList.SelectedItem, ProjFileInfo)
        Dim fileName = projFileInfo.FilePath

        If File.Exists(fileName) Then
            Try
                FileIO.FileSystem.DeleteFile(fileName, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin)
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If

        fileName = fileName.Substring(0, fileName.Length - 4) & "sb"
        If File.Exists(fileName) Then
            Try
                FileIO.FileSystem.DeleteFile(fileName, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin)
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If

        fileName &= ".gen"
        If File.Exists(fileName) Then
            Try
                FileIO.FileSystem.DeleteFile(fileName, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin)
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If

        projFiles.Remove(projFileInfo)
        RaiseEvent FileDeleted()

        Dim n = FilesList.Items.Count
        If n = 0 Then Return
        FilesList.SelectedIndex = If(i >= n, n - 1, i)
    End Sub

    Protected Overrides Function OnCommit(newName As String) As Boolean
        Dim item = CType(FilesList.SelectedItem, ProjFileInfo)
        If item.IsTheGlobalFile Then Return False

        Dim oldFile = item.FilePath
        Dim x = newName(0).ToString().ToUpper()
        newName = x & If(newName.Length > 1, newName.Substring(1), "")
        Dim newFile = Path.Combine(_projDir, newName & ".xaml")

        Try
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


                Dim oldGenFile = oldFile & ".gen"
                Dim newGenFile = newFile & ".gen"
                If File.Exists(oldGenFile) AndAlso Not File.Exists(newGenFile) Then
                    File.Move(oldGenFile, newGenFile)
                End If

                RaiseEvent FileNameChanged(oldFile, newFile)

                Return True
            Else
                MsgBox($"The file `{oldFile}` doesn't exist in project folder", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical)
                Return False
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        Return False
    End Function

    Protected Overrides Function OnBeginEdit() As Boolean
        Return Not CType(FilesList.SelectedItem, ProjFileInfo).IsTheGlobalFile
    End Function

    Public Sub SelectedGlobalFile()
        SelectItem(Path.Combine(_projDir, "Global.sb"))
    End Sub
End Class

Public Class ProjFileInfo

    Dim _filePath As String

    Public Sub New(filePath As String)
        _filePath = filePath
        IsTheGlobalFile = Path.GetFileNameWithoutExtension(filePath).ToLower() = "global"
    End Sub

    Public Property FilePath As String
        Get
            Return _filePath
        End Get

        Set(value As String)
            _filePath = value
        End Set
    End Property

    Public ReadOnly Property IsTheGlobalFile As Boolean


    Public Overrides Function ToString() As String
        Return " ● " & Path.GetFileName(_filePath)
    End Function
End Class