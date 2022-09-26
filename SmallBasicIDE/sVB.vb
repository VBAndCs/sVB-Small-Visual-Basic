Imports Microsoft.SmallVisualBasic
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

    Private Sub New(genCode As String, code As String)
        ' A new parser for each code file
        Compiler.CreateNewParser()
        hints = WinForms.PreCompiler.ParseFormHints(genCode)
        AddCodeLines(genCode)
        LineOffset = If(lines.Count = 0, 0, lines.Count)
        AddCodeLines(code)

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


    Public Shared Function Compile(genCode As String, code As String) As List(Of [Error])
        Dim sVBCompiler As New sVB(genCode, code)
        With sVBCompiler
            .errors = .Compile()
            If .errors.Count > 0 Then
                Dim controlNamesList As New Dictionary(Of String, String)
                Dim variableTypes As New Dictionary(Of String, VariableType)
                Dim normalErrors As New List(Of [Error])

                For Each [error] In .errors
                    Dim errMsg = [error].Description
                    If Not (errMsg = "!" OrElse errMsg.StartsWith("Cannot find object", StringComparison.InvariantCultureIgnoreCase)) Then
                        normalErrors.Add([error])

                    ElseIf normalErrors.Count = 0 Then
                        Dim pos = errMsg.LastIndexOf("'", errMsg.Length - 3) + 1
                        Dim obj As String = errMsg.Substring(pos, errMsg.Length - pos - 2).ToLower()

                        If Not controlNamesList.ContainsKey(obj) AndAlso Not variableTypes.ContainsKey(obj) Then
                            If .hints IsNot Nothing AndAlso .hints.ControlsInfo.ContainsKey(obj) Then
                                controlNamesList(obj) = .hints.ControlsInfo(obj)
                            Else
                                Dim controlName = WinForms.PreCompiler.GetModuleFromVarName(obj)
                                If controlName = "" Then
                                    Dim varType = sVB._compiler.Parser.SymbolTable.GetInferedType([error].Token)
                                    If varType = VariableType.None Then
                                        normalErrors.Add([error])
                                    Else
                                        variableTypes(obj) = varType
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

                .PreCompile(controlNamesList, variableTypes)
            End If

            Return .errors
        End With
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
                    If Not controlNames.ContainsKey(obj) AndAlso Not variableTypes.ContainsKey(obj) Then
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

    Private Function Compile() As List(Of [Error])
        If hints IsNot Nothing Then
            _compiler.Parser.SymbolTable.ModuleNames = hints.ControlsInfo
            _compiler.Parser.SymbolTable.ControlNames = hints.ControlNames
        End If
        Return _compiler.Compile(lines)
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

            Dim controlName = If(controlNames.ContainsKey(obj), controlNames(obj), "")
            Dim varType = If(variableTypes.ContainsKey(obj), variableTypes(obj), VariableType.None)

            'use (lineNum) to be passed by value
            Dim tokens = LineScanner.GetTokens(line, (lineNum), lines)

            Dim objId As Integer
            For objId = 0 To tokens.Count - 1
                If tokens(objId).Line = lineNum AndAlso tokens(objId).Column = charNum Then Exit For
            Next

            If objId + 2 >= tokens.Count Then
                errors(i) = New [Error](lineNum, tokens(objId + 1).EndColumn, charNum, "Method name is expected.")
                Continue For
            End If

            If tokens(objId + 1).Type <> TokenType.Dot Then Continue For

            Dim nameToken = tokens(objId + 2)
            Dim prevText = If(charNum = 0, "", line.Substring(0, charNum))
            Dim methodPos = charNum + obj.Length + 1
            Dim nextText = line.Substring(methodPos)

            If objId + 3 < tokens.Count AndAlso nameToken.Type = TokenType.Identifier AndAlso tokens(objId + 3).Type = TokenType.LeftParens Then
                ' Method Call
                Dim method = nameToken.Text
                Dim restText = line.Substring(tokens(objId + 3).Column + 1)
                Dim argsExprList = Parser.ParseArgumentList(restText, lineNum, lines, TokenType.LeftParens)
                If argsExprList Is Nothing Then
                    errors(i) = New [Error](nameToken, "Wrong brackets pairs")
                    Continue For
                End If

                Dim methodInfo = WinForms.PreCompiler.GetMethodInfo(controlName, varType, method)
                Dim ModuleName = methodInfo.Module
                If ModuleName = "" Then
                    errors(i) = New [Error](nameToken, $"Method `{method}` doesn't exist.")
                    Continue For
                End If

                Dim argsCount = argsExprList.Count
                Dim paramsCount = methodInfo.ParamsCount
                If paramsCount = 0 OrElse paramsCount <= argsCount OrElse
                          argsCount <> paramsCount - 1 Then
                    errors(i) = New [Error](nameToken, "Wrong number of parameters.")
                    Continue For
                End If

                lines(lineNum) = prevText &
                        $"{ModuleName}.{method}({objName}, " &
                        restText

                errors.RemoveAt(i)
                ReRun = True

            ElseIf objId = 0 AndAlso err.subLine = 0 Then 'Property Set 

                If nameToken.Type <> TokenType.Identifier Then
                    errors(i) = New [Error](nameToken, $"Expected a property name.")
                    Continue For
                End If

                If objId + 3 >= tokens.Count OrElse tokens(objId + 3).Type <> TokenType.Equals Then
                    errors(i) = New [Error](nameToken, $"Expected `=` and a value to set the property")
                    Continue For
                End If

                Dim propName = nameToken.Text

                Dim propInfo = WinForms.PreCompiler.GetMethodInfo(controlName, varType, propName)
                If propInfo.Module = "" Then
                    Dim method = $"Set{propName}"
                    Dim methodInfo = WinForms.PreCompiler.GetMethodInfo(controlName, varType, method)
                    Dim ModuleName = methodInfo.Module

                    If ModuleName = "" Then
                        errors(i) = New [Error](nameToken, $"Property `{propName}` doesn't exist.")
                        Continue For
                    End If

                    If methodInfo.ParamsCount <> 2 Then
                        errors(i) = New [Error](nameToken, $"`{method}` definition is not supported.")
                        Continue For
                    End If

                    pos = tokens(objId + 3).Column
                    lines(lineNum) = prevText &
                            $"{ModuleName}.{method}({obj}, {line.Substring(pos + 1).Trim}"
                    lines(lineNum + tokens.Last.subLine) += ")"
                    errors.RemoveAt(i)
                    ReRun = True

                Else ' Event
                    Dim ModuleName = propInfo.Module
                    lines.Insert(lineNum, $"Control.HandleEvents({obj})")
                    lines(lineNum + 1) = prevText & $"{ModuleName}.{nextText}"
                    errors.RemoveAt(i)
                    ReRun = True
                End If

            Else 'Property Get                     
                If tokens.Count > objId + 2 AndAlso nameToken.Type <> TokenType.Identifier Then
                    errors(i) = New [Error](nameToken, "property name is expected.")
                    Continue For
                End If

                Dim propName = nameToken.Text
                Dim method = $"Get{propName}"
                Dim methodInfo = WinForms.PreCompiler.GetMethodInfo(controlName, varType, method)
                Dim ModuleName = methodInfo.Module

                If ModuleName = "" Then
                    errors(i) = New [Error](nameToken, $"Property `{propName}` doesn't exist.")
                    Continue For
                End If

                If methodInfo.ParamsCount = 1 Then
                    lines(lineNum) =
                       prevText &
                       $"{ModuleName}.{method}({obj})" &
                       nextText.Substring(propName.Length)
                Else
                    errors(i) = New [Error](nameToken, $"`{method}` definition is not supported.")
                    Continue For
                End If

                errors.RemoveAt(i)
                ReRun = True
            End If

        Next

        If ReRun Then Compile(
            controlNames,
            variableTypes
        )
    End Sub

End Class
