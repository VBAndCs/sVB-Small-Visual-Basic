﻿Imports Microsoft.SmallVisualBasic
Imports Microsoft.VisualBasic
Imports System
Imports System.Collections.Generic
Imports System.Linq

Class sVB
    Private ReadOnly lines As New List(Of String)
    Private ReadOnly hints As WinForms.FormInfo
    Dim errors As List(Of [Error])
    Private Shared _compiler As Compiler
    Public Shared LineOffset As Integer

    Private Sub New(genCode As String, code As String, formNames As List(Of String))
        ' A new parser for each code file
        Compiler.CreateNewParser()
        _compiler.Parser.FormNames = formNames

        If genCode = "" Then
            LineOffset = 0
        Else
            hints = WinForms.PreCompiler.ParseFormHints(genCode)
            AddCodeLines(genCode)
            LineOffset = If(lines.Count = 0, 0, lines.Count)
        End If

        AddCodeLines(code)
        _compiler.Parser.DocStartLine = LineOffset
        If ClassName <> "" Then
            _compiler.Parser.ClassName = ClassName
        End If
    End Sub

    Sub AddCodeLines(code As String)
        Dim start As Integer = 0
        Dim n = code.Length - 1
        Dim i = 0

        Do Until i > n
            If code(i) = vbCr Then
                lines.Add(code.Substring(start, i - start))
                If i = n Then Exit Do
                i += If(code(i + 1) = vbLf, 2, 1)
                start = i

            ElseIf code(i) = vbLf Then
                lines.Add(code.Substring(start, i - start))
                i += 1
                start = i

            Else
                i += 1
            End If
        Loop

        If start <= n Then
            lines.Add(code.Substring(start, n - start + 1))
        End If
    End Sub

    Public Shared ReadOnly Property Compiler As Compiler
        Get
            If _compiler Is Nothing Then _compiler = New Compiler()
            Return _compiler
        End Get
    End Property

    Private Shared _mainWindow As MainWindow
    Friend Shared ClassName As String

    Private Shared ReadOnly Property MainWindow As MainWindow
        Get
            If _mainWindow Is Nothing Then
                _mainWindow = Windows.Application.Current.MainWindow
            End If
            Return _mainWindow
        End Get
    End Property

    Public Shared Function GetOutputFileName(filePath As String, isSingleCodeFile As Boolean) As String
        If filePath = "" Then
            Dim tempFolder = Now.ToString("yy-MM-dd-HH-mm-ss", New Globalization.CultureInfo("en-US"))
            tempFolder = IO.Path.Combine(App.SvbTempFolder, tempFolder)
            If Not IO.Directory.Exists(tempFolder) Then
                IO.Directory.CreateDirectory(tempFolder)
            End If
            Return IO.Path.Combine(tempFolder, "untitled.exe").ToLower()
        End If

        Dim docDirectory = IO.Path.GetDirectoryName(filePath)
        Dim fileName = IO.Path.GetFileNameWithoutExtension(If(isSingleCodeFile, filePath, docDirectory))
        If fileName = "" Then
            fileName = IO.Path.GetFileNameWithoutExtension(filePath)
        End If

        Dim binDirectory = IO.Path.Combine(docDirectory, "bin")
        If Not IO.Directory.Exists(binDirectory) Then
            IO.Directory.CreateDirectory(binDirectory)
        End If
        Dim newFile = IO.Path.Combine(binDirectory, fileName)

        For Each f In IO.Directory.EnumerateFiles(docDirectory)
            Select Case IO.Path.GetExtension(f).ToLower().TrimStart("."c)
                Case "bmp", "jpg", "jpeg", "png", "gif", "ico",
                     "wav", "mp3", "mpa", "wma",
                     "txt", "xaml", "style"
                    Dim f2 = IO.Path.Combine(binDirectory, IO.Path.GetFileName(f))
                    Try
                        IO.File.Copy(f, f2, True)
                    Catch
                    End Try
            End Select
        Next
        Return newFile & ".exe"
    End Function

    Public Shared Function CompileGlobalModule(
                   inputDir As String,
                   outputFileName As String,
                  Optional formNames As List(Of String) = Nothing,
                   Optional ignoreErrors As Boolean = True
               ) As List(Of Parser)

        Dim globalFile = IO.Path.Combine(inputDir, "global.sb")
        Dim code = ""
        Dim parsers As New List(Of Parser)

        If IO.File.Exists(globalFile) Then
            _compiler = New Compiler()
            Compiler.ExeFile = outputFileName
            Dim globalDoc = MainWindow.GetDocIfOpened(globalFile)
            Dim compileGlobal = True

            If globalDoc Is Nothing Then
                If ignoreErrors AndAlso Compiler.GlobalParser IsNot Nothing Then
                    parsers.Add(Compiler.GlobalParser)
                    compileGlobal = False
                Else
                    code = IO.File.ReadAllText(globalFile)
                End If

            ElseIf globalDoc.IsDirty Then
                If Not ignoreErrors Then globalDoc.Save()
                If globalDoc.LastModified > Compiler.GlobalLastCompiled Then
                    code = globalDoc.Text
                Else
                    parsers.Add(Compiler.GlobalParser)
                    compileGlobal = False
                End If

            ElseIf ignoreErrors AndAlso Compiler.GlobalParser IsNot Nothing Then
                parsers.Add(Compiler.GlobalParser)
                compileGlobal = False
            Else
                code = globalDoc.Text
            End If

            If compileGlobal Then
                Dim errors = Compile("", code, True, ignoreErrors, formNames)

                If errors?.Count = 0 Then
                    parsers.Add(Compiler.Parser)

                ElseIf ignoreErrors Then
                    Return Nothing

                Else
                    If globalDoc Is Nothing Then
                        globalDoc = MainWindow.OpenDocIfNot(globalFile)
                    End If

                    globalDoc.ShowErrors(errors)
                    MainWindow.tabCode.IsSelected = True
                    Return Nothing
                End If
            End If
        End If

        If parsers.Count > 0 Then
            parsers(0).DocPath = globalFile
        End If

        ' Note that return nothing means there are errors in code
        ' but returning empty list of parsers mean ther is no global file
        Return parsers
    End Function

    Public Shared Function Compile(
                     genCode As String,
                     code As String,
                     doc As Documents.TextDocument,
                     parsers As List(Of Parser)
                ) As Boolean


        Dim errors = Compile(genCode, code)

        If errors.Count = 0 Then
            Compiler.Parser.DocPath = doc.File
            parsers.Add(Compiler.Parser)
            Return True
        Else
            doc.ShowErrors(errors)
            MainWindow.tabCode.IsSelected = True
            Return False
        End If
    End Function

    Public Shared Function Compile(
                    genCode As String,
                    code As String,
                    Optional isGlobal As Boolean = False,
                    Optional ignoreErrors As Boolean = False,
                    Optional formNames As List(Of String) = Nothing
               ) As List(Of [Error])

        Dim sVBCompiler As New sVB(genCode, code, formNames)
        _compiler.Parser.IsGlobal = isGlobal

        sVBCompiler.errors = sVBCompiler.Compile(ignoreErrors)
        If sVBCompiler.errors?.Count > 0 Then
            Dim controlNamesList As New Dictionary(Of String, String)
            Dim variableTypes As New Dictionary(Of String, VariableType)
            Dim normalErrors As New List(Of [Error])

            For Each [error] In sVBCompiler.errors
                Dim errMsg = [error].Description
                If Not (errMsg = "!" OrElse errMsg.StartsWith("Cannot find object", StringComparison.InvariantCultureIgnoreCase)) Then
                    normalErrors.Add([error])

                ElseIf normalErrors.Count = 0 Then
                    Dim pos = errMsg.LastIndexOf("'", errMsg.Length - 3) + 1
                    Dim obj As String = errMsg.Substring(pos, errMsg.Length - pos - 2).ToLower()
                    Dim key = obj

                    If _compiler.Parser.SymbolTable.IsLocalVar([error].Token) Then
                        key = [error].Token.SubroutineName.ToLower() + "." + obj
                    End If

                    If Not controlNamesList.ContainsKey(obj) AndAlso Not variableTypes.ContainsKey(key) Then
                        If sVBCompiler.hints IsNot Nothing AndAlso sVBCompiler.hints.ControlsInfo.ContainsKey(obj) Then
                            controlNamesList(obj) = sVBCompiler.hints.ControlsInfo(obj)
                        Else
                            Dim controlName = WinForms.PreCompiler.GetModuleFromVarName(obj)
                            If controlName = "" Then
                                Dim varType = _compiler.Parser.SymbolTable.GetInferedType([error].Token)
                                If varType = VariableType.Any Then
                                    variableTypes(key) = VariableType.String
                                Else
                                    variableTypes(key) = varType
                                End If
                            Else
                                controlNamesList(obj) = controlName
                            End If
                        End If
                    End If
                End If
            Next

            If normalErrors.Count > 0 Then
                ' Fix normal errors first
                Return normalErrors
            End If

            sVBCompiler.PreCompile(controlNamesList, variableTypes)
        End If

        Return sVBCompiler.errors
    End Function

    Private Sub Compile(
                     controlNames As Dictionary(Of String, String),
                     variableTypes As Dictionary(Of String, VariableType)
                )

        errors = Compile()

        If errors.Count > 0 Then
            Dim normalErrors As New List(Of [Error])

            For Each [error] In errors
                Dim errMsg = [error].Description
                If Not (errMsg = "!" OrElse errMsg.StartsWith("Cannot find object", StringComparison.InvariantCultureIgnoreCase)) Then
                    normalErrors.Add([error])
                ElseIf normalErrors.Count = 0 Then
                    Dim pos = errMsg.LastIndexOf("'", errMsg.Length - 3) + 1
                    Dim obj As String = errMsg.Substring(pos, errMsg.Length - pos - 2).ToLower()
                    Dim key = obj

                    If _compiler.Parser.SymbolTable.IsLocalVar([error].Token) Then
                        key = [error].Token.SubroutineName.ToLower() + "." + obj
                    End If

                    If Not controlNames.ContainsKey(obj) AndAlso Not variableTypes.ContainsKey(key) Then
                        normalErrors.Add([error])
                    End If
                End If
            Next

            If normalErrors.Count > 0 Then
                ' Fix normal errors first
                errors = normalErrors
                Return
            End If

            PreCompile(controlNames, variableTypes)
        End If

    End Sub

    Private Function Compile(Optional ignoreErrors As Boolean = False) As List(Of [Error])
        If hints IsNot Nothing Then
            _compiler.Parser.SymbolTable.ModuleNames = hints.ControlsInfo
            _compiler.Parser.SymbolTable.ControlNames = hints.ControlNames
        End If
        Return _compiler.Compile(lines, ignoreErrors)
    End Function

    Private Sub PreCompile(
                   controlNames As Dictionary(Of String, String),
                   variableTypes As Dictionary(Of String, VariableType)
             )

        Dim ReRun = False

        Dim lineNum As Integer
        For i = errors.Count - 1 To 0 Step -1
            Dim err = errors(i)
            If lineNum = err.Line AndAlso ReRun Then Exit For

            Dim errMsg = err.Description
            lineNum = err.Line
            Dim charNum = err.Column
            Dim line = lines(lineNum)

            If errMsg = "!" Then
                lines(lineNum) = line.Substring(0, charNum) & "Data." & line.Substring(charNum + 1)
                errors.RemoveAt(i)
                ReRun = True
                Continue For
            End If

            Dim pos = errMsg.LastIndexOf("'", errMsg.Length - 3) + 1
            Dim objName = errMsg.Substring(pos, errMsg.Length - pos - 2)
            Dim obj = objName.ToLower()
            Dim key = obj

            If _compiler.Parser.SymbolTable.IsLocalVar(err.Token) Then
                key = err.Token.SubroutineName.ToLower() + "." + obj
            End If

            Dim controlName = If(controlNames.ContainsKey(obj), controlNames(obj), "")
            Dim varType = If(variableTypes.ContainsKey(key), variableTypes(key), VariableType.Any)

            'use (lineNum) to be passed by value
            Dim tokens = LineScanner.GetTokens(line, (lineNum), lines)
            Dim objId As Integer

            For objId = 0 To tokens.Count - 1
                If tokens(objId).Line = lineNum AndAlso tokens(objId).Column = charNum Then Exit For
            Next

            If objId + 2 >= tokens.Count Then
                errors.Clear()
                errors.Add(New [Error](lineNum, tokens(objId + 1).EndColumn, charNum, "Method name is expected."))
                Exit Sub
            End If

            If tokens(objId + 1).Type <> TokenType.Dot Then Continue For

            Dim nameToken = tokens(objId + 2)
            Dim prevText = If(charNum = 0, "", line.Substring(0, charNum))
            Dim methodPos = charNum + obj.Length + 1
            Dim lastToken = tokens.Last
            Dim commentPos = If(
                lastToken.Line = lineNum AndAlso lastToken.Type = TokenType.Comment,
                lastToken.Column,
                -1
            )

            Dim nextText = line.Substring(
                methodPos,
                If(commentPos = -1, line.Length, commentPos) - methodPos
            )

            If objId + 3 < tokens.Count AndAlso nameToken.Type = TokenType.Identifier AndAlso tokens(objId + 3).Type = TokenType.LeftParens Then
                ' Method Call
                Dim method = nameToken.Text
                Dim restText = line.Substring(tokens(objId + 3).Column + 1)
                Dim argsExprList = Parser.ParseArgumentList(restText, lineNum, lines, TokenType.LeftParens)

                If argsExprList Is Nothing Then
                    errors.Clear()
                    errors.Add(New [Error](nameToken, "Wrong brackets pairs"))
                    Exit Sub
                End If

                Dim methodInfo = WinForms.PreCompiler.GetMethodInfo(controlName, varType, method)
                Dim ModuleName = methodInfo.Module
                If ModuleName = "" Then
                    errors.Clear()
                    errors.Add(New [Error](nameToken, $"Method `{method}` doesn't exist."))
                    Exit Sub
                End If

                Dim argsCount = argsExprList.Count
                Dim paramsCount = methodInfo.ParamsCount
                If paramsCount = 0 OrElse paramsCount <= argsCount OrElse
                          argsCount <> paramsCount - 1 Then
                    errors.Clear()
                    errors.Add(New [Error](nameToken, "Wrong number of parameters."))
                    Exit Sub
                End If

                lines(lineNum) = prevText &
                        $"{ModuleName}.{method}({objName}{If(argsCount = 0, "", ", ")}" &
                        restText

                errors.RemoveAt(i)
                ReRun = True

            ElseIf objId = 0 AndAlso err.subLine = 0 Then 'Property Set 
                If nameToken.Type <> TokenType.Identifier Then
                    errors.Clear()
                    errors.Add(New [Error](nameToken, $"Expected a property name."))
                    Exit Sub
                End If

                If objId + 3 >= tokens.Count OrElse tokens(objId + 3).Type <> TokenType.EqualsTo Then
                    errors.Clear()
                    errors.Add(New [Error](nameToken, $"Expected `=` And a value to set the property"))
                    Exit Sub
                End If

                Dim propName = nameToken.Text

                Dim propInfo = WinForms.PreCompiler.GetMethodInfo(controlName, varType, propName)
                If propInfo.Module = "" Then
                    Dim method = $"Set{propName}"
                    Dim methodInfo = WinForms.PreCompiler.GetMethodInfo(controlName, varType, method)
                    Dim ModuleName = methodInfo.Module

                    If ModuleName = "" Then
                        errors.Clear()
                        methodInfo = WinForms.PreCompiler.GetMethodInfo(controlName, varType, $"Get{propName}")
                        If methodInfo.Module = "" Then
                            errors.Add(New [Error](nameToken, $"The `{propName}` property doesn't exist."))
                        Else
                            errors.Add(New [Error](nameToken, $"The `{propName}` property is read-only, and you can't change its value."))
                        End If
                        Exit Sub
                    End If

                    If methodInfo.ParamsCount <> 2 Then
                        errors.Clear()
                        errors.Add(New [Error](nameToken, $"The `{method}` definition is not supported."))
                        Exit Sub
                    End If

                    pos = tokens(objId + 3).Column
                    lines(lineNum) = prevText &
                            $"{ModuleName}.{method}({obj}, {line.Substring(pos + 1).Trim}"

                    Dim lastLineNum = lineNum + lastToken.subLine
                    If commentPos = -1 AndAlso lastToken.Type <> TokenType.Comment Then
                        lines(lastLineNum) += ")"
                    Else
                        Dim x = lines(lastLineNum)
                        lines(lastLineNum) = x.Substring(0, x.Length - lastToken.EndColumn + lastToken.Column) & ")" & lastToken.Text
                    End If

                    errors.RemoveAt(i)
                    ReRun = True

                Else ' Event
                    Dim ModuleName = propInfo.Module
                    If tokens(tokens.Count - 1).Type = TokenType.Nothing Then
                        Dim t = _compiler.Parser.SymbolTable.GetInferedType(obj)
                        If t = VariableType.WinTimer Then
                            lines(lineNum) = $"WinTimer.RemoveOnTickHandler({obj})"
                        ElseIf t = VariableType.Form AndAlso tokens(2).Text.ToLower() = "onclosed" Then
                            lines(lineNum) = $"Control.HandleEvents({obj})" & vbLf &
                                                        prevText & $"{ModuleName}.{nextText}"
                        Else
                            lines(lineNum) = $"Control.RemoveEventHandler({obj}, ""{tokens(2).Text}"")"
                        End If
                    Else
                        lines(lineNum) = $"Control.HandleEvents({obj})" & vbLf &
                                                   prevText & $"{ModuleName}.{nextText}"
                    End If
                    errors.RemoveAt(i)
                    ReRun = True
                End If

            Else 'Property Get                     
                If tokens.Count > objId + 2 AndAlso nameToken.Type <> TokenType.Identifier Then
                    errors.Clear()
                    errors.Add(New [Error](nameToken, "property name is expected."))
                    Exit Sub
                End If

                Dim propName = nameToken.Text
                Dim method = $"Get{propName}"
                Dim methodInfo = WinForms.PreCompiler.GetMethodInfo(controlName, varType, method)
                Dim ModuleName = methodInfo.Module

                If ModuleName = "" Then
                    errors.Clear()
                    errors.Add(New [Error](nameToken, $"Property `{propName}` doesn't exist."))
                    Exit Sub
                End If

                If methodInfo.ParamsCount = 1 Then
                    lines(lineNum) =
                       prevText &
                       $"{ModuleName}.{method}({obj})" &
                       nextText.Substring(propName.Length) &
                       If(commentPos > -1, lastToken.Text, "")
                Else
                    errors.Clear()
                    errors.Add(New [Error](nameToken, $"`{method}` definition is not supported."))
                    Exit Sub
                End If

                errors.RemoveAt(i)
                ReRun = True
            End If

        Next

        If ReRun Then Compile(controlNames, variableTypes)
    End Sub

End Class
