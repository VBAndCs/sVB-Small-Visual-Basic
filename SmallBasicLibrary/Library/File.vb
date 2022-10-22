
Imports System.IO
Imports System.Text
Imports Microsoft.SmallVisualBasic.Library.Internal

Namespace Library
    ''' <summary>
    ''' The File object provides methods to access, read and write information from and to a file on disk.  Using this object, it is possible to save and open settings across multiple sessions of your program.
    ''' </summary>
    <SmallBasicType>
    Public NotInheritable Class File
        ''' <summary>
        ''' <para>
        ''' Gets or sets the last encountered file operation based error message.  This property is useful for finding out when some method fails to execute.
        ''' </para>
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Property LastError As Primitive

        ''' <summary>
        ''' <para>
        ''' Opens a file and reads the entire file's contents.  This method will be fast for small files that are less than an MB in size, but will start to slow down and will be noticeable for files greater than 10MB.
        ''' </para>
        ''' </summary>
        ''' <param name="filePath">
        ''' <para>
        ''' The full path of the file to read.  An example of a full path will be c:\temp\settings.data.
        ''' </para>
        ''' </param>
        ''' <returns>
        ''' The entire contents of the file.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function ReadContents(filePath As Primitive) As Primitive
            LastError = ""
            Dim path As String = Environment.ExpandEnvironmentVariables(filePath)
            Try
                If Not IO.File.Exists(path) Then
                    Return ""
                End If

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
        ''' <para>
        ''' Opens a file and writes the specified contents into it, replacing the original contents with the new content.
        ''' </para>
        ''' </summary>
        ''' <param name="filePath">
        ''' <para>
        ''' The full path of the file to write to.  An example of a full path will be c:\temp\settings.data.
        ''' </para>
        ''' </param>
        ''' <param name="contents">
        ''' The contents to write into the specified file.
        ''' </param>
        ''' <returns>
        ''' <para>
        ''' If the operation was successful, this will return "SUCCESS".  Otherwise, it will return "FAILED".
        ''' </para>
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function WriteContents(filePath As Primitive, contents As Primitive) As Primitive
            Try
                LastError = ""
                Dim path As String = Environment.ExpandEnvironmentVariables(filePath)
                Using writer As New StreamWriter(path, False, Encoding.Default)
                    writer.Write(CStr(contents))
                    Return "SUCCESS"
                End Using
            Catch ex As Exception
                LastError = ex.Message
            End Try
            Return "FAILED"
        End Function

        ''' <summary>
        ''' Opens the specified file and reads the contents at the specified line number.
        ''' </summary>
        ''' <param name="filePath">
        ''' <para>
        ''' The full path of the file to read from.  An example of a full path will be c:\temp\settings.data.
        ''' </para>
        ''' </param>
        ''' <param name="lineNumber">
        ''' The line number of the text to be read.
        ''' </param>
        ''' <returns>
        ''' The text at the specified line of the specified file.
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function ReadLine(filePath As Primitive, lineNumber As Primitive) As Primitive
            LastError = ""
            Dim path As String = Environment.ExpandEnvironmentVariables(filePath)
            If CBool((lineNumber < 0)) Then
                Return ""
            End If
            Try
                If Not IO.File.Exists(path) Then
                    Return ""
                End If
                Using streamReader1 As New StreamReader(path)
                    For i As Integer = 0 To CInt(lineNumber) - 2
                        If streamReader1.ReadLine() Is Nothing Then
                            Return ""
                        End If
                    Next

                    Return streamReader1.ReadLine()
                End Using
            Catch ex As Exception
                LastError = ex.Message
            End Try

            Return ""
        End Function

        ''' <summary>
        ''' <para>
        ''' Opens the specified file and write the contents at the specified line number.
        ''' </para>
        ''' <para>
        ''' This operation will overwrite any existing content at the specified line.
        ''' </para>
        ''' </summary>
        ''' <param name="filePath">
        ''' <para>
        ''' The full path of the file to read from.  An example of a full path will be c:\temp\settings.data.
        ''' </para>
        ''' </param>
        ''' <param name="lineNumber">
        ''' The line number of the text to write.
        ''' </param>
        ''' <param name="contents">
        ''' <para>
        ''' The contents to write at the specified line of the specified file.
        ''' </para>
        ''' </param>
        ''' <returns>
        ''' <para>
        ''' If the operation was successful, this will return "SUCCESS".  Otherwise, it will return "FAILED".
        ''' </para>
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function WriteLine(filePath As Primitive, lineNumber As Primitive, contents As Primitive) As Primitive
            LastError = ""
            Dim path1 As String = Environment.ExpandEnvironmentVariables(filePath)
            Dim tempFileName As String = Path.GetTempFileName()
            Try
                If Not IO.File.Exists(path1) Then
                    Using streamWriter1 As New StreamWriter(path1)
                        streamWriter1.WriteLine(CStr(contents))
                    End Using

                    Return "SUCCESS"
                End If
                Using streamWriter2 As New StreamWriter(tempFileName)
                    Using streamReader1 As New StreamReader(path1)
                        Dim num As Integer = 1
                        While True
                            Dim text As String = streamReader1.ReadLine()
                            If text Is Nothing Then
                                Exit While
                            End If
                            If CBool((num = lineNumber)) Then
                                streamWriter2.WriteLine(CStr(contents))
                            Else
                                streamWriter2.WriteLine(text)
                            End If
                            num += 1
                        End While
                        If CBool((num <= lineNumber)) Then
                            streamWriter2.WriteLine(CStr(contents))
                        End If
                    End Using
                End Using
                IO.File.Copy(tempFileName, filePath, overwrite:=True)
                IO.File.Delete(tempFileName)
                Return "SUCCESS"
            Catch ex As Exception
                LastError = ex.Message
            End Try

            Return "FAILED"
        End Function

        ''' <summary>
        ''' <para>
        ''' Opens the specified file and inserts the contents at the specified line number.
        ''' </para>
        ''' <para>
        ''' This operation will not overwrite any existing content at the specified line.
        ''' </para>
        ''' </summary>
        ''' <param name="filePath">
        ''' <para>
        ''' The full path of the file to read from.  An example of a full path will be c:\temp\settings.data.
        ''' </para>
        ''' </param>
        ''' <param name="lineNumber">
        ''' The line number of the text to insert.
        ''' </param>
        ''' <param name="contents">
        ''' The contents to insert into the file.
        ''' </param>
        ''' <returns>
        ''' <para>
        ''' If the operation was successful, this will return "SUCCESS".  Otherwise, it will return "FAILED".
        ''' </para>
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function InsertLine(filePath As Primitive, lineNumber As Primitive, contents As Primitive) As Primitive
            LastError = ""
            Dim path1 As String = Environment.ExpandEnvironmentVariables(filePath)
            Dim tempFileName As String = Path.GetTempFileName()
            Try
                If Not IO.File.Exists(path1) Then
                    Using streamWriter1 As New StreamWriter(path1)
                        streamWriter1.WriteLine(CStr(contents))
                    End Using

                    Return "SUCCESS"
                End If
                Using streamWriter2 As New StreamWriter(tempFileName)
                    Using streamReader1 As New StreamReader(path1)
                        Dim num As Integer = 1
                        While True
                            Dim text As String = streamReader1.ReadLine()
                            If text Is Nothing Then
                                Exit While
                            End If
                            If CBool((num = lineNumber)) Then
                                streamWriter2.WriteLine(CStr(contents))
                            End If
                            streamWriter2.WriteLine(text)
                            num += 1
                        End While
                        If CBool((num <= lineNumber)) Then
                            streamWriter2.WriteLine(CStr(contents))
                        End If
                    End Using
                End Using
                IO.File.Copy(tempFileName, filePath, overwrite:=True)
                IO.File.Delete(tempFileName)
                Return "SUCCESS"
            Catch ex As Exception
                LastError = ex.Message
            End Try

            Return "FAILED"
        End Function

        ''' <summary>
        ''' <para>
        ''' Opens the specified file and appends the contents to the end of the file.
        ''' </para>
        ''' </summary>
        ''' <param name="filePath">
        ''' <para>
        ''' The full path of the file to read from.  An example of a full path will be c:\temp\settings.data.
        ''' </para>
        ''' </param>
        ''' <param name="contents">
        ''' The contents to append to the end of the file.
        ''' </param>
        ''' <returns>
        ''' <para>
        ''' If the operation was successful, this will return "SUCCESS".  Otherwise, it will return "FAILED".
        ''' </para>
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function AppendContents(filePath As Primitive, contents As Primitive) As Primitive
            LastError = ""
            Environment.ExpandEnvironmentVariables(filePath)
            Try
                Using streamWriter1 As New StreamWriter(filePath, append:=True)
                    streamWriter1.WriteLine(CStr(contents))
                End Using

                Return "SUCCESS"
            Catch ex As Exception
                LastError = ex.Message
            End Try

            Return "FAILED"
        End Function

        ''' <summary>
        ''' <para>
        ''' Copies the specified source file to the destination file path.  If the destination points to a location that doesn't exist, the method will attempt to create it automatically.
        ''' </para>
        ''' <para>
        ''' Existing files will be overwritten.  It is always best to check if the destination file exists if you don't want to overwrite existing files.
        ''' </para>
        ''' </summary>
        ''' <param name="sourceFilePath">
        ''' <para>
        ''' The full path of the file that needs to be copied.  An example of a full path will be c:\temp\settings.data.
        ''' </para>
        ''' </param>
        ''' <param name="destinationFilePath">
        ''' The destination location or the file path.  
        ''' </param>
        ''' <returns>
        ''' If the operation was successful, this will return "SUCCESS".  Otherwise, it will return "FAILED".
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function CopyFile(sourceFilePath As Primitive, destinationFilePath As Primitive) As Primitive
            LastError = ""
            Dim text As String = Environment.ExpandEnvironmentVariables(sourceFilePath)
            Dim text2 As String = Environment.ExpandEnvironmentVariables(destinationFilePath)
            If Not IO.File.Exists(text) Then
                LastError = "Source file doesn't exist."
                Return "FAILED"
            End If
            If Directory.Exists(text2) OrElse text2(text2.Length - 1) = "\"c Then
                text2 = Path.Combine(text2, Path.GetFileName(text))
            End If
            Try
                Dim directoryName As String = Path.GetDirectoryName(text2)
                If Not Directory.Exists(directoryName) Then
                    Directory.CreateDirectory(directoryName)
                End If
                IO.File.Copy(text, text2, overwrite:=True)
                Return "SUCCESS"
            Catch ex As Exception
                LastError = ex.Message
            End Try

            Return "FAILED"
        End Function

        ''' <summary>
        ''' Deletes the specified file.
        ''' </summary>
        ''' <param name="filePath">
        ''' <para>
        ''' The destination location or the file path.  An example of a full path will be
        ''' c:\temp\settings.data.
        ''' </para>
        ''' </param>
        ''' <returns>
        ''' If the operation was successful, this will return "SUCCESS".  Otherwise, it will return "FAILED".
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function DeleteFile(filePath As Primitive) As Primitive
            LastError = ""
            Dim path As String = Environment.ExpandEnvironmentVariables(filePath)
            Try
                IO.File.Delete(path)
                Return "SUCCESS"
            Catch ex As Exception
                LastError = ex.Message
            End Try

            Return "FAILED"
        End Function

        ''' <summary>
        ''' Creates the specified directory.
        ''' </summary>
        ''' <param name="directoryPath">
        ''' The full path of the directory to be created.
        ''' </param>
        ''' <returns>
        ''' If the operation was successful, this will return "SUCCESS".  Otherwise, it will return "FAILED".
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function CreateDirectory(directoryPath As Primitive) As Primitive
            LastError = ""
            Environment.ExpandEnvironmentVariables(directoryPath)
            Try
                Directory.CreateDirectory(directoryPath)
                Return "SUCCESS"
            Catch ex As Exception
                LastError = ex.Message
            End Try

            Return "FAILED"
        End Function

        ''' <summary>
        ''' Deletes the specified directory.
        ''' </summary>
        ''' <param name="directoryPath">
        ''' The full path of the directory to be deleted.
        ''' </param>
        ''' <returns>
        ''' If the operation was successful, this will return "SUCCESS".  Otherwise, it will return "FAILED".
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared Function DeleteDirectory(directoryPath As Primitive) As Primitive
            LastError = ""
            Environment.ExpandEnvironmentVariables(directoryPath)
            Try
                Directory.Delete(directoryPath, recursive:=True)
                Return "SUCCESS"
            Catch ex As Exception
                LastError = ex.Message
            End Try

            Return "FAILED"
        End Function

        ''' <summary>
        ''' Gets the path of all the files in the specified directory path.
        ''' </summary>
        ''' <param name="directoryPath">
        ''' The directory to look for files.
        ''' </param>
        ''' <returns>
        ''' If the operation was successful, this will return the files as an array.  Otherwise, it will return "FAILED".
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.Array)>
        Public Shared Function GetFiles(directoryPath As Primitive) As Primitive
            LastError = ""
            Environment.ExpandEnvironmentVariables(directoryPath)
            Try
                If Directory.Exists(directoryPath) Then
                    Dim dictionary1 As New Dictionary(Of Primitive, Primitive)
                    Dim num As Integer = 1
                    Dim files As String() = Directory.GetFiles(directoryPath)
                    For Each value As String In files
                        dictionary1(num) = value
                        num += 1
                    Next

                    Return Primitive.ConvertFromMap(dictionary1)
                End If
                LastError = "Directory Path does not exist."
                Return "FAILED"
            Catch ex As Exception
                LastError = ex.Message
                Return "FAILED"
            End Try
        End Function

        ''' <summary>
        ''' Gets the path of all the directories in the specified directory path.
        ''' </summary>
        ''' <param name="directoryPath">
        ''' The directory to look for subdirectories.
        ''' </param>
        ''' <returns>
        ''' If the operation was successful, this will return the list of directories as an array.  Otherwise, it will return "FAILED".
        ''' </returns>
        <WinForms.ReturnValueType(VariableType.Array)>
        Public Shared Function GetDirectories(directoryPath As Primitive) As Primitive
            LastError = ""
            Environment.ExpandEnvironmentVariables(directoryPath)
            Try
                If Directory.Exists(directoryPath) Then
                    Dim dictionary1 As New Dictionary(Of Primitive, Primitive)
                    Dim num As Integer = 1
                    Dim directories As String() = Directory.GetDirectories(directoryPath)
                    For Each value As String In directories
                        dictionary1(num) = value
                        num += 1
                    Next

                    Return Primitive.ConvertFromMap(dictionary1)
                End If
                LastError = "Directory Path does not exist."
                Return "FAILED"
            Catch ex As Exception
                LastError = ex.Message
                Return "FAILED"
            End Try
        End Function

        ''' <summary>
        ''' <para>
        ''' Creates a new temporary file in a temporary directory and returns the 
        ''' full file path.
        ''' </para>
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
            Dim entryAssemblyPath As String = SmallBasicApplication.GetEntryAssemblyPath()
            Return Path.ChangeExtension(entryAssemblyPath, ".settings")
        End Function

        ''' <summary>
        ''' Shows the open file dialog, to allow the user to select a file from his PC file system.
        ''' </summary>
        ''' <param name="extFilters">
        ''' An array containing extention filters, to allow you specify the file types you want to open. Each extenstion filter itself is an array, where the first Item descrips the file type, and the rset of items contains one or more extebrion, such as {"Images", "bmp", "jpg", "gif"}
        ''' If you will open only one file type catagory like images, you can sent the extention filter directly like {"Images", "bmp", "jpg", "gif"}.
        ''' else, send an array of extenstion filter like: {{"Text Files", ".txt"}, {"Images", "bmp", "jpg", "gif"}}
        ''' If you will open only a single extenstion, you can just use it as a single string like: "doc". This can also be done for extention filter elements, such as:
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
                .FilterIndex = GetSetting("sVB", "OpenFile", key & "_FilterIndex", 1)
            }

            If dlg.ShowDialog() = True Then
                SaveSetting("sVB", "OpenFile", key, IO.Path.GetDirectoryName(dlg.FileName))
                SaveSetting("sVB", "OpenFile", key & "_FilterIndex", dlg.FilterIndex)
                Return dlg.FileName
            Else
                Return ""
            End If
        End Function

        ''' <summary>
        ''' Shows the save file dialog, to allow the user to choose a folder from his PC file system to save the file to.
        ''' </summary>
        ''' <param name="fileName">The name to save the file with. User can change this name in the dialog. You can use the full path of the file, to suggest the initial directory in the dialog, otherwise, the initial directory will be the last opened one</param>
        ''' <param name="extFilters">
        ''' An array containing extention filters, to allow you specify the file types you want to save. Each extenstion filter itself is an array, where the first Item descrips the file type, and the rset of items contains one or more extebrion, such as {"Images", "bmp", "jpg", "gif"}
        ''' If you will open only one file type catagory like images, you can sent the extention filter directly like {"Images", "bmp", "jpg", "gif"}.
        ''' else, send an array of extenstion filter like: {{"Text Files", ".txt"}, {"Images", "bmp", "jpg", "gif"}}
        ''' If you will open only a single extenstion, you can just use it as a single string like: "doc". This can also be done for extention filter elements, such as:
        ''' {{"Text Files", ".txt"}, {"Images", "bmp", "jpg", "gif"}, "doc"}
        ''' If you want to show all filess, use "" or "*" as the extension, like: {"All Files", "*"}, or just use "*".
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
                    Dim x = Str.Substring(0, pos)
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

        Private Shared Function GetFilter(fileType As String)
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
