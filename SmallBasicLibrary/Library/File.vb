
Imports System.IO
Imports System.Text
Imports Microsoft.SmallVisualBasic.Library.Internal

Namespace Library

    ''' <summary>
    ''' The File object provides methods to access, read and write information from and to a file on disk.  Using this object, it is possible to save and open settings across multiple sessions of your program.
    ''' </summary>
    <SmallVisualBasicType>
    Public NotInheritable Class File

        ''' <summary>
        ''' Gets or sets the last encountered file operation based error message.  This property is useful for finding out when some method fails to execute.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Property LastError As Primitive

        ''' <summary>
        ''' Opens a file and reads the entire file's contents.  This method will be fast for small files that are less than an MB in size, but will start to slow down and will be noticeable for files greater than 10MB.
        ''' </summary>
        ''' <param name="filePath">
        ''' The full path of the file to read, like c:\temp\settings.data.
        ''' </param>
        ''' <returns>The entire contents of the file.</returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function ReadContents(filePath As Primitive) As Primitive
            LastError = ""
            Dim path As String = Environment.ExpandEnvironmentVariables(filePath)
            Try
                If Not IO.File.Exists(path) Then Return ""

                Using reader As New StreamReader(path, Encoding.Default)
                    Dim value As String = reader.ReadToEnd()
                    Return value
                End Using

            Catch ex As Exception
                LastError = ex.Message
            End Try

            Return ""
        End Function

        ''' <summary>
        ''' Opens a file and writes the specified contents into it, replacing the original contents with the new content.
        ''' </summary>
        ''' <param name="filePath">The full path of the file to write to, like c:\temp\settings.data.</param>
        ''' <param name="contents">The contents to write into the specified file. You can send an array to write its elements to the file, each element in a line</param>
        ''' <returns>True if the operation was successful, or False otherwise.</returns>
        <WinForms.ReturnValueType(VariableType.Boolean)>
        Public Shared Function WriteContents(filePath As Primitive, contents As Primitive) As Primitive
            Try
                LastError = ""
                Dim path As String = Environment.ExpandEnvironmentVariables(filePath)

                Using writer As New StreamWriter(path, False, Encoding.Default)
                    If contents.IsArray Then
                        Dim lines = contents._arrayMap.Values
                        Dim n = lines.Count - 1
                        For i = 0 To n - 1
                            writer.WriteLine(lines(i).AsString())
                        Next
                        writer.Write(lines(n).AsString())
                    Else
                        writer.Write(contents.AsString())
                    End If
                End Using

                Return True

            Catch ex As Exception
                LastError = ex.Message
            End Try

            Return False
        End Function

        ''' <summary>
        ''' Opens the specified file and reads the contents at the specified line number.
        ''' Don't use this method in a loop to read many lines from the same file, because it aleays starts reading from the first line of the file until it reaches the rewuired line, which will have a very bad impact in loops.
        ''' In such case, you can use the ReadLines method to read all the file lines then walk through it using the loop as you want.
        ''' </summary>
        ''' <param name="filePath">The full path of the file to read from, like c:\temp\settings.data.</param>
        ''' <param name="lineNumber">The line number of the text to be read.</param>
        ''' <returns>The text at the specified line of the specified file.</returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function ReadLine(filePath As Primitive, lineNumber As Primitive) As Primitive
            LastError = ""
            Dim path As String = Environment.ExpandEnvironmentVariables(filePath)
            If lineNumber < 1 Then Return ""

            Try
                If Not IO.File.Exists(path) Then Return ""

                Using reader As New StreamReader(path, Encoding.Default)
                    For i = 1 To CInt(lineNumber) - 1
                        If reader.ReadLine() Is Nothing Then Return ""
                    Next

                    Return reader.ReadLine()
                End Using
            Catch ex As Exception
                LastError = ex.Message
            End Try

            Return ""
        End Function

        ''' <summary>
        ''' Opens the specified file and reads all its lines.
        ''' </summary>
        ''' <param name="filePath">The full path of the file to read from, like c:\temp\settings.data.</param>
        ''' <returns>an array containing the lines of the specified file.</returns>
        <WinForms.ReturnValueType(VariableType.Array)>
        Public Shared Function ReadLines(filePath As Primitive) As Primitive
            LastError = ""
            Dim path As String = Environment.ExpandEnvironmentVariables(filePath)

            Try
                If Not IO.File.Exists(path) Then Return ""
                Dim lines As New Dictionary(Of Primitive, Primitive)
                Dim n = 1

                Using reader As New StreamReader(path, Encoding.Default)
                    Do
                        Dim line = reader.ReadLine()
                        If line Is Nothing Then Exit Do
                        lines(n) = line
                        n += 1
                    Loop
                End Using

                Dim result As New Primitive
                result._arrayMap = lines
                Return result

            Catch ex As Exception
                LastError = ex.Message
            End Try

            Return ""
        End Function


        ''' <summary>
        ''' Opens the specified file and writes the contents at the specified line number.
        ''' This operation will overwrite any existing content at the specified line.
        ''' Don't use this method to write many lines in the same file using a loop, as it will have a bad impact on performance, because it involves copying all file lines to a temp file to insert the the given content at the given line number, which will be repeated for every line you write using the loop!
        ''' In such case, use the AppendContents method insread, to add accumulated content to the end of the file, or add all the lines to one variable and use the InsertLine method to insert it for once.
        ''' </summary>
        ''' <param name="filePath">The full path of the file to read from.  An example of a full path will be c:\temp\settings.data.</param>
        ''' <param name="lineNumber">The line number of the text to write.</param>
        ''' <param name="contents">The contents to write at the specified line of the specified file.</param>
        ''' <returns>True if the operation was successful, or False otherwise.</returns>
        <WinForms.ReturnValueType(VariableType.Boolean)>
        Public Shared Function WriteLine(
                         filePath As Primitive,
                         lineNumber As Primitive,
                         contents As Primitive
                   ) As Primitive

            LastError = ""
            Dim path1 = Environment.ExpandEnvironmentVariables(filePath)
            Dim tempFileName = Path.GetTempFileName()

            Try
                If Not IO.File.Exists(path1) Then
                    Using writer1 As New StreamWriter(path1, False, Encoding.Default)
                        writer1.WriteLine(CStr(contents))
                    End Using

                    Return True
                End If

                Using writer2 As New StreamWriter(tempFileName, False, Encoding.Default)
                    Using reader As New StreamReader(path1, Encoding.Default)
                        Dim num As Integer = 1
                        While True
                            Dim text As String = reader.ReadLine()
                            If text Is Nothing Then Exit While

                            If num = lineNumber Then
                                writer2.WriteLine(CStr(contents))
                            Else
                                writer2.WriteLine(text)
                            End If
                            num += 1
                        End While

                        If num <= lineNumber Then
                            writer2.WriteLine(CStr(contents))
                        End If
                    End Using
                End Using

                IO.File.Copy(tempFileName, filePath, overwrite:=True)
                IO.File.Delete(tempFileName)
                Return True

            Catch ex As Exception
                LastError = ex.Message
            End Try

            Return False
        End Function

        ''' <summary>
        ''' Opens the specified file and inserts the contents at the specified line number.
        ''' This operation will not overwrite any existing content at the specified line.
        ''' Don't use this method to write many lines in the same file using a loop, as it will have a bad impact on performance, because it involves copying all file lines to a temp file to insert the the given content at the given line number, which will be reoeated for every line you write using the loop!
        ''' In such case, use the AppendContents method insread, to add accumulated content to the end of the file, or add all the lines to one variable and use the InsertLine method to insert it for once.
        ''' </summary>
        ''' <param name="filePath">The full path of the file to read from, like c:\temp\settings.data.</param>
        ''' <param name="lineNumber">The line number of the text to insert.</param>
        ''' <param name="contents">The contents to insert into the file.</param>
        ''' <returns>True if the operation was successful, or False otherwise.</returns>
        <WinForms.ReturnValueType(VariableType.Boolean)>
        Public Shared Function InsertLine(filePath As Primitive, lineNumber As Primitive, contents As Primitive) As Primitive
            LastError = ""
            Dim path1 = Environment.ExpandEnvironmentVariables(filePath)
            Dim tempFileName = Path.GetTempFileName()

            Try
                If Not IO.File.Exists(path1) Then
                    Using writer1 As New StreamWriter(path1, False, Encoding.Default)
                        writer1.WriteLine(CStr(contents))
                    End Using

                    Return True
                End If

                Using writer2 As New StreamWriter(tempFileName, False, Encoding.Default)
                    Using reader As New StreamReader(path1, Encoding.Default)
                        Dim num As Integer = 1
                        While True
                            Dim text As String = reader.ReadLine()
                            If text Is Nothing Then Exit While

                            If num = lineNumber Then
                                writer2.WriteLine(CStr(contents))
                            End If

                            writer2.WriteLine(text)
                            num += 1
                        End While

                        If num <= lineNumber Then
                            writer2.WriteLine(CStr(contents))
                        End If
                    End Using
                End Using

                IO.File.Copy(tempFileName, filePath, overwrite:=True)
                IO.File.Delete(tempFileName)
                Return True

            Catch ex As Exception
                LastError = ex.Message
            End Try

            Return False
        End Function

        ''' <summary>
        ''' Opens the specified file and appends the contents to the end of the file then adds a new line.
        ''' </summary>
        ''' <param name="filePath">The full path of the file to read from.  An example of a full path will be c:\temp\settings.data.</param>
        ''' <param name="contents">The contents to append to the end of the file.</param>
        ''' <returns>True if the operation was successful, or False otherwise.</returns>
        <WinForms.ReturnValueType(VariableType.Boolean)>
        <HideFromIntellisense>
        Public Shared Function AppendContents(filePath As Primitive, contents As Primitive) As Primitive
            Return AppendLine(filePath, contents)
        End Function

        ''' <summary>
        ''' Opens the specified file and appends the contents to the end of the file then adds a new line.
        ''' </summary>
        ''' <param name="filePath">The full path of the file to write to.  An example of a full path will be c:\temp\settings.data.</param>
        ''' <param name="contents">The contents to append to the end of the file. You can send an array to append its elements to the file, each element in a line.</param>
        ''' <returns>True if the operation was successful, or False otherwise.</returns>
        <WinForms.ReturnValueType(VariableType.Boolean)>
        <HideFromIntellisense>
        Public Shared Function AppendLine(filePath As Primitive, contents As Primitive) As Primitive
            LastError = ""
            Dim path1 = Environment.ExpandEnvironmentVariables(filePath)

            Try
                Using writer As New StreamWriter(path1, append:=True, Encoding.Default)
                    If contents.IsArray Then
                        For Each line In contents._arrayMap.Values
                            writer.WriteLine(line.AsString())
                        Next
                    Else
                        writer.WriteLine(contents.AsString())
                    End If
                End Using

                Return True

            Catch ex As Exception
                LastError = ex.Message
            End Try

            Return False
        End Function

        ''' <summary>
        ''' Opens the specified file and appends the contents to the end of the file.
        ''' </summary>
        ''' <param name="filePath">The full path of the file to read from.  An example of a full path will be c:\temp\settings.data.</param>
        ''' <param name="contents">The contents to append to the end of the file.</param>
        ''' <returns>True if the operation was successful, or False otherwise.</returns>
        <WinForms.ReturnValueType(VariableType.Boolean)>
        <HideFromIntellisense>
        Public Shared Function Append(filePath As Primitive, contents As Primitive) As Primitive
            LastError = ""
            Dim path1 = Environment.ExpandEnvironmentVariables(filePath)

            Try
                Using writer As New StreamWriter(path1, append:=True, Encoding.Default)
                    writer.Write(contents.AsString())
                End Using
                Return True

            Catch ex As Exception
                LastError = ex.Message
            End Try

            Return False
        End Function


        ''' <summary>
        ''' Copies the specified source file to the destination file path.  If the destination points to a location that doesn't exist, the method will attempt to create it automatically.
        ''' Existing files will be overwritten. It is always best to check if the destination file exists if you don't want to overwrite existing files.
        ''' </summary>
        ''' <param name="sourceFilePath">The full path of the file that needs to be copied.  An example of a full path will be c:\temp\settings.data.</param>
        ''' <param name="destinationFilePath">The destination location or the file path.</param>
        ''' <returns>True if the operation was successful, or False otherwise.</returns>
        <WinForms.ReturnValueType(VariableType.Boolean)>
        Public Shared Function CopyFile(sourceFilePath As Primitive, destinationFilePath As Primitive) As Primitive
            LastError = ""
            Dim file1 = Environment.ExpandEnvironmentVariables(sourceFilePath)
            Dim file2 = Environment.ExpandEnvironmentVariables(destinationFilePath)

            If Not IO.File.Exists(file1) Then
                LastError = "Source file doesn't exist."
                Return False
            End If

            If Directory.Exists(file2) OrElse file2.EndsWith("\") Then
                file2 = Path.Combine(file2, Path.GetFileName(file1))
            End If

            Try
                Dim directoryName = Path.GetDirectoryName(file2)
                If Not Directory.Exists(directoryName) Then
                    Directory.CreateDirectory(directoryName)
                End If

                IO.File.Copy(file1, file2, overwrite:=True)
                Return True

            Catch ex As Exception
                LastError = ex.Message
            End Try

            Return False
        End Function

        ''' <summary>
        ''' Deletes the specified file.
        ''' </summary>
        ''' <param name="filePath">The destination location or the file path, like c:\temp\settings.data.</param>
        ''' <returns>True if the operation was successful, or False otherwise.</returns>
        <WinForms.ReturnValueType(VariableType.Boolean)>
        Public Shared Function DeleteFile(filePath As Primitive) As Primitive
            LastError = ""
            Dim path As String = Environment.ExpandEnvironmentVariables(filePath)

            Try
                IO.File.Delete(path)
                Return True
            Catch ex As Exception
                LastError = ex.Message
            End Try

            Return False
        End Function

        ''' <summary>
        ''' Creates the specified directory.
        ''' </summary>
        ''' <param name="directoryPath">The full path of the directory to be created.</param>
        ''' <returns>True if the operation was successful, or False otherwise.</returns>
        <WinForms.ReturnValueType(VariableType.Boolean)>
        Public Shared Function CreateDirectory(directoryPath As Primitive) As Primitive
            LastError = ""
            Dim dir = Environment.ExpandEnvironmentVariables(directoryPath)

            Try
                Directory.CreateDirectory(dir)
                Return True
            Catch ex As Exception
                LastError = ex.Message
            End Try

            Return False
        End Function

        ''' <summary>
        ''' Deletes the specified directory.
        ''' </summary>
        ''' <param name="directoryPath">The full path of the directory to be deleted.</param>
        ''' <returns>True if the operation was successful, or False otherwise.</returns>
        <WinForms.ReturnValueType(VariableType.Boolean)>
        Public Shared Function DeleteDirectory(directoryPath As Primitive) As Primitive
            LastError = ""
            Dim dir = Environment.ExpandEnvironmentVariables(directoryPath)

            Try
                Directory.Delete(dir, recursive:=True)
                Return True
            Catch ex As Exception
                LastError = ex.Message
            End Try

            Return False
        End Function

        ''' <summary>
        ''' Gets the path of all the files in the specified directory path.
        ''' </summary>
        ''' <param name="directoryPath">The directory to look for files.</param>
        ''' <returns>
        ''' If the operation was successful, this will return the files as an array.  Otherwise, it will return an empty string "".
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.Array)>
        Public Shared Function GetFiles(directoryPath As Primitive) As Primitive
            LastError = ""
            Dim dir = Environment.ExpandEnvironmentVariables(directoryPath)

            Try
                If Directory.Exists(dir) Then
                    Dim dictionary As New Dictionary(Of Primitive, Primitive)
                    Dim n As Integer = 1
                    Dim files = Directory.GetFiles(dir)

                    For Each file In files
                        dictionary(n) = file
                        n += 1
                    Next

                    Return Primitive.ConvertFromMap(dictionary)
                End If

                LastError = "Directory Path does not exist."
                Return ""

            Catch ex As Exception
                LastError = ex.Message
                Return ""
            End Try
        End Function

        ''' <summary>
        ''' Gets the path of all the directories in the specified directory path.
        ''' </summary>
        ''' <param name="directoryPath">The directory to look for subdirectories.</param>
        ''' <returns>
        ''' If the operation was successful, this will return the list of directories as an array.  Otherwise, it will return an empty string "".
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.Array)>
        Public Shared Function GetDirectories(directoryPath As Primitive) As Primitive
            LastError = ""
            Dim dir = Environment.ExpandEnvironmentVariables(directoryPath)

            Try
                If Directory.Exists(dir) Then
                    Dim dictionary As New Dictionary(Of Primitive, Primitive)
                    Dim n As Integer = 1
                    Dim directories = Directory.GetDirectories(dir)

                    For Each item In directories
                        dictionary(n) = item
                        n += 1
                    Next

                    Return Primitive.ConvertFromMap(dictionary)
                End If

                LastError = "Directory Path does not exist."
                Return ""

            Catch ex As Exception
                LastError = ex.Message
                Return ""
            End Try
        End Function

        ''' <summary>
        ''' Creates a new temporary file in a temporary directory and returns the full file path.
        ''' </summary>
        ''' <returns>
        ''' The full file path of the temporary file.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function GetTemporaryFilePath() As Primitive
            Return Path.GetTempFileName()
        End Function

        ''' <summary>
        ''' Gets the full path of the settings file for this program.  The settings file name is based on the program's name and is present in the same location as the program.
        ''' </summary>
        ''' <returns>
        ''' The full path of the settings file specific for this program.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function GetSettingsFilePath() As Primitive
            Dim entryAssemblyPath = SmallBasicApplication.GetEntryAssemblyPath()
            Return Path.ChangeExtension(entryAssemblyPath, ".settings")
        End Function

        ''' <summary>
        ''' Shows the open file dialog, to allow the user to select a file from his PC file system.
        ''' </summary>
        ''' <param name="extFilters">
        ''' An array containing extension filters, to allow you specify the file types you want to open. Each extenstion filter itself is an array, where the first Item descries the file type, and the rset of items contains one or more extebrion, such as {"Images", "bmp", "jpg", "gif"}
        ''' If you will open only one file type catagory like images, you can sent the extension filter directly like {"Images", "bmp", "jpg", "gif"}.
        ''' else, send an array of extenstion filter like: {{"Text Files", ".txt"}, {"Images", "bmp", "jpg", "gif"}}
        ''' If you will open only a single extenstion, you can just use it as a single string like: "doc". This can also be done for extension filter elements, such as:
        ''' {{"Text Files", ".txt"}, {"Images", "bmp", "jpg", "gif"}, "doc"}
        ''' If you want to show all filess, use "" or "*" as the extension, like: {"All Files", "*"}, or just use "*".
        ''' </param>
        ''' <returns>
        ''' The file name that the user selected, or an empty string "" if he canceled the operation.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function OpenFileDialog(extFilters As Primitive) As Primitive
            Dim filter = BuildFilter(extFilters)
            Dim key = GetKey(filter)

            Dim dlg As New Win32.OpenFileDialog With {
                .Filter = filter,
                .Title = "Open File",
                .RestoreDirectory = False,
                .InitialDirectory = GetSetting("sVB", "OpenFile", key, ""),
                .FilterIndex = CInt(GetSetting("sVB", "OpenFile", key & "_FilterIndex", "1"))
            }

            If dlg.ShowDialog() = True Then
                SaveSetting("sVB", "OpenFile", key, IO.Path.GetDirectoryName(dlg.FileName))
                SaveSetting("sVB", "OpenFile", key & "_FilterIndex", dlg.FilterIndex.ToString())
                Return dlg.FileName
            Else
                Return ""
            End If
        End Function

        ''' <summary>
        ''' Shows the open folder dialog, to allow the user to select a folder from his PC file system.
        ''' </summary>
        ''' <param name="initialFolder">
        ''' The folder path to be initialy selected when the dialog is shown.
        ''' Use an empty string to let the dialog show the last folder that was previously selected by the user.
        ''' </param>
        ''' <returns>
        ''' The folder name that the user selected, or an empty string "" if he canceled the operation.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function OpenFolderDialog(initialFolder As Primitive) As Primitive
            SmallBasicApplication.Invoke(
                Sub()
                    Dim folder = GeIinitialFolder(initialFolder.AsString())
                    If folder = "" Then folder = GeIinitialFolder(System.Windows.Clipboard.GetText())
                    If folder = "" Then folder = GetSetting("sVB", "OpenFolder", "LastFolder", "")

                    Dim th As New Threading.Thread(
                        Sub()
                            Threading.Thread.Sleep(300)
                            System.Windows.Forms.SendKeys.SendWait("{TAB}{TAB}{RIGHT}")
                        End Sub)
                    th.Start()


                    Try
                        Dim dlg As New System.Windows.Forms.FolderBrowserDialog With {
                            .Description = "Select a folder:",
                            .SelectedPath = folder
                        }

                        If dlg.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                            SaveSetting("sVB", "OpenFolder", "LastFolder", dlg.SelectedPath)
                            OpenFolderDialog = dlg.SelectedPath
                        End If

                    Catch ex As Exception
                        MsgBox(ex.Message)
                    End Try
                End Sub)

        End Function

        Private Shared Function GeIinitialFolder(folder As String) As String
            Try
                If folder <> "" Then
                    If IO.File.Exists(folder) Then
                        folder = Path.GetDirectoryName(folder)
                    Else
                        Do
                            If Directory.GetDirectoryRoot(folder) = folder OrElse Directory.Exists(folder) Then
                                Exit Do
                            End If
                            folder = Path.GetDirectoryName(folder)
                        Loop
                    End If
                End If

            Catch
                folder = ""
            End Try

            Return folder
        End Function

        ''' <summary>
        ''' Shows the save file dialog, to allow the user to choose a folder from his PC file system to save the file to.
        ''' </summary>
        ''' <param name="fileName">The name to save the file with. User can change this name in the dialog. You can use the full path of the file, to suggest the initial directory in the dialog, otherwise, the initial directory will be the last opened one</param>
        ''' <param name="extFilters">
        ''' An array containing extension filters, to allow you specify the file types you want to save. Each extenstion filter itself is an array, where the first Item descries the file type, and the rset of items contains one or more extebrion, such as {"Images", "bmp", "jpg", "gif"}
        ''' If you will open only one file type catagory like images, you can sent the extension filter directly like {"Images", "bmp", "jpg", "gif"}.
        ''' else, send an array of extenstion filter like: {{"Text Files", ".txt"}, {"Images", "bmp", "jpg", "gif"}}
        ''' If you will open only a single extenstion, you can just use it as a single string like: "doc". This can also be done for extension filter elements, such as:
        ''' {{"Text Files", ".txt"}, {"Images", "bmp", "jpg", "gif"}, "doc"}
        ''' If you want to show all files, use "" or "*" as the extension, like: {"All Files", "*"}, or just use "*".
        ''' In fact, you can combine all filters in on string using the .NET standard formula, where you separate filter parts with | and separate extensions with ; like:
        ''' "Images|*.bmp;*.jpg;*.gif|Text Files|*.txt|docs|*.doc|All Files|*.*"
        ''' </param>
        ''' <returns>
        ''' The file name that the user selected, or an empty string "" if he canceled the operation.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function SaveFileDialog(fileName As Primitive, extFilters As Primitive) As Primitive
            Dim filter = BuildFilter(extFilters)
            Dim key = GetKey(filter)
            Dim d As String
            Dim file = fileName.ToString()
            Dim name = ""
            Dim index As Primitive

            If file <> "" Then
                name = IO.Path.GetFileName(file)

                If file = "" OrElse file.ToLower() = name.ToLower() Then
                    d = GetSetting("sVB", "SaveFile", key, "")
                Else
                    d = IO.Path.GetDirectoryName(file)
                End If

                Dim ext = IO.Path.GetExtension(name).TrimStart("."c)
                If extFilters.IsArray Then
                    index = Array.Find(extFilters, ext, 1, True)
                Else
                    Dim str = extFilters.AsString()
                    Dim pos = str.LastIndexOf("*." & ext)
                    Dim x = str.Substring(0, pos)
                    index = (x.Length - x.Replace("|", "").Length + 1) \ 2
                End If

                If index.IsEmpty Then
                    For Each item In extFilters._arrayMap.Values
                        index = Array.Find(item, ext, 1, True)
                        If Not index.IsEmpty Then Exit For
                    Next
                End If
            End If

            Dim dlg As New Win32.SaveFileDialog With {
                .Filter = filter,
                .Title = "Save File",
                .RestoreDirectory = False,
                .InitialDirectory = d,
                .FileName = name,
                .FilterIndex = index
            }

            If dlg.ShowDialog() = True Then
                SaveSetting("sVB", "SaveFile", key, IO.Path.GetDirectoryName(dlg.FileName))
                Return dlg.FileName
            Else
                Return ""
            End If
        End Function

        Private Shared Function GetKey(filter As String) As String
            Return filter.Replace("|", "_").
                                 Replace(";", "_").
                                 Replace(",", "_").
                                 Replace(" ", "_").
                                 Replace("(", "_").
                                 Replace(")", "_").
                                 Replace("*", "_").
                                 Replace(".", "_")
        End Function

        Private Shared Function GetFilter(fileType As String) As String
            If fileType.Contains("|") OrElse fileType.Contains(";") Then
                Return fileType
            Else
                Dim x = fileType.Trim("."c, "*"c)
                If x = "" Then x = "*"
                Dim ext = "*." & x
                Return $"{ext}|{ext}"
            End If
        End Function

        Private Shared Function BuildFilter(extFilters As Primitive) As String
            Dim filter As New StringBuilder
            If extFilters.IsArray Then
                If extFilters._arrayMap("1").IsArray Then
                    For Each ext In extFilters._arrayMap.Values
                        If ext.IsArray Then
                            AddFileType(filter, ext)
                        Else
                            filter.Append(GetFilter(ext))
                            filter.Append("|")
                        End If
                    Next
                Else
                    AddFileType(filter, extFilters)
                End If
            Else
                filter.Append(GetFilter(extFilters))
            End If
            Return filter.ToString().TrimEnd("|"c)

        End Function

        Private Shared Sub AddFileType(filter As StringBuilder, ext As Primitive)
            Dim st = 0
            Dim parts = ext._arrayMap.Values
            Dim description = parts(0).AsString()
            If Not description.Contains(".") Then
                st = 1
                filter.Append(description)
                filter.Append("|")
            End If

            Dim n = parts.Count - 1
            For i = st To n
                filter.Append("*.")
                Dim x = parts(i).AsString().Trim("*"c, "."c)
                If x = "" Then x = "*"
                filter.Append(x)
                filter.Append(If(i < n, ";", "|"))
            Next
        End Sub
    End Class
End Namespace
