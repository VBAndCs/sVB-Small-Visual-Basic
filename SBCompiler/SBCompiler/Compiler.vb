Imports System.IO
Imports System.Reflection
Imports Microsoft.SmallBasic
Imports Microsoft.SmallVisualBasic.Library

Namespace Microsoft.SmallVisualBasic
    Public Class Compiler
        Private Shared _referenceAssemblies As New List(Of Assembly)
        Private Shared _libraryFiles As New List(Of String)()
        Private Shared _typeInfoBag As TypeInfoBag
        Private Shared _references As New List(Of String)

        Private _errors As New List(Of [Error])()

        Public ReadOnly Property References As List(Of String)
            Get
                Return _references
            End Get
        End Property

        Public ReadOnly Property Parser As Parser
        Public Property GlobalParser As Parser
        Public GlobalLastCompiled As Date = Date.MinValue
        Dim lastModified As Date = Date.MinValue
        Dim _exeFile As String

        Public Property ExeFile As String
            Get
                Return _exeFile
            End Get

            Set(value As String)
                value = value.ToLower()
                If _exeFile = value Then
                    Dim binFolder = Path.GetDirectoryName(_exeFile)
                    Dim codeFolder = Path.GetDirectoryName(binFolder)
                    Dim globalFile = Path.Combine(codeFolder, "global.sb")

                    Dim d = If(
                        IO.File.Exists(globalFile),
                        IO.File.GetLastWriteTime(globalFile),
                        Date.MinValue
                    )

                    If _GlobalParser Is Nothing Then
                        lastModified = d
                    ElseIf d > lastModified Then
                        lastModified = d
                        _GlobalParser = Nothing
                        _typeInfoBag.Types.Remove("global")
                    End If

                Else
                    _exeFile = value
                    lastModified = Now
                    _GlobalParser = Nothing
                    _typeInfoBag.Types.Remove("global")
                End If
            End Set
        End Property

        Public Shared ReadOnly Property TypeInfoBag As TypeInfoBag
            Get
                Return _typeInfoBag
            End Get
        End Property

        Shared Sub New()
            _typeInfoBag = New TypeInfoBag()
            Initialize()
        End Sub

        Public Sub New()
            _Parser = New Parser(_errors)
        End Sub

        Public Sub CreateNewParser()
            _Parser = New Parser(_errors)
        End Sub

        Private Shared Sub Initialize()
            PopulateReferences()
            PopulateClrSymbols()
            PopulatePrimitiveMethods()
        End Sub

        Public Function Compile(
                        source As TextReader,
                        Optional autoCompletion As Boolean = False
                   ) As List(Of [Error])

            If source Is Nothing Then
                Throw New ArgumentNullException("source")
            End If

            Dim codeLines As New List(Of String)

            Do
                Dim line = source.ReadLine()
                If line Is Nothing Then Exit Do
                codeLines.Add(line)
            Loop

            Return Compile(codeLines, autoCompletion)
        End Function

        Public Function Compile(
                        codeLines As List(Of String),
                        Optional ignoreErrors As Boolean = False
                    ) As List(Of [Error])

            _errors.Clear()
            _Parser.Parse(
                codeLines,
                ignoreErrors,
                If(ignoreErrors, _GlobalParser, Nothing)
            )

            If Not ignoreErrors Then
                If _errors.Count > 0 Then Return _errors
                Dim analyzer As New SemanticAnalyzer(_Parser, _typeInfoBag)
                analyzer.Analyze()
                If _errors.Count > 0 Then Return _errors
            End If

            ' Generate in memory Global type. We need this even at build time,
            ' because other forms may use the global type,
            ' so, we need it to compile them correctly
            If _Parser.IsGlobal Then
                _Parser.ClassName = "Global"
                Dim codeGenerator As New CodeGenerator(
                        New List(Of Parser) From {Parser},
                        _typeInfoBag,
                        If(_exeFile = "", "tempGlobal", Path.GetFileNameWithoutExtension(_exeFile)),
                        If(_exeFile = "", svbDir, Path.GetDirectoryName(_exeFile))
                )
                codeGenerator.GenerateExecutable(True) ' Aleays use try not to emit to exe
                _GlobalParser = _Parser
                GlobalLastCompiled = Now
            End If

            Return _errors

        End Function

        Public Sub Build(
                       parsers As List(Of Parser),
                       exeFile As String
                   )

            If parsers Is Nothing OrElse parsers.Count = 0 Then
                Throw New ArgumentNullException("parsers")
            End If

            If exeFile = "" Then
                Throw New ArgumentNullException("exeFile")
            End If

            Dim outputName = Path.GetFileNameWithoutExtension(exeFile)
            Dim directory As String = Path.GetDirectoryName(exeFile)

            Build(parsers, outputName, directory)
        End Sub

        Public Function Build(
                        parsers As List(Of Parser),
                        outputName As String,
                        directory As String
                   ) As List(Of [Error])

            If parsers Is Nothing OrElse parsers.Count = 0 Then
                Throw New ArgumentNullException("parsers")
            End If

            If Equals(outputName, Nothing) Then
                Throw New ArgumentNullException("outputName")
            End If

            If Equals(directory, Nothing) Then
                Throw New ArgumentNullException("directory")
            End If

            Dim gen As New CodeGenerator(
                    parsers, _typeInfoBag, outputName, directory)
            Dim refs = gen.GenerateExecutable()
            CopyLibraryAssemblies(directory, refs)
            Return _errors
        End Function

        Private Sub CopyLibraryAssemblies(
                    directory As String,
                    refs() As AssemblyName
            )

            Dim location = GetType(Primitive).Assembly.Location
            Dim fileName = Path.GetFileName(location)

            Try
                IO.File.Copy(location, Path.Combine(directory, fileName), overwrite:=True)
            Catch
            End Try

            Dim copySBLib = False

            For Each libraryFile In _libraryFiles
                Try
                    Dim asmName = Path.GetFileNameWithoutExtension(libraryFile)
                    Dim found = False
                    For Each ref In refs
                        If ref.Name.ToLower() = asmName.ToLower() Then
                            found = True
                            Exit For
                        End If
                    Next
                    If Not found Then Continue For

                    fileName = Path.GetFileName(libraryFile)
                    If fileName.ToLower().EndsWith(".exe") Then
                        Dim dllName = fileName.Substring(0, fileName.Length - 3) & "dll"
                        IO.File.Copy(libraryFile, Path.Combine(directory, dllName))

                        Dim dir = IO.Path.GetDirectoryName(libraryFile)
                        For Each file In IO.Directory.GetFiles(dir)
                            Select Case IO.Path.GetExtension(file).ToLower().TrimStart("."c)
                                Case "bmp", "jpg", "jpeg", "png", "gif", "txt", "xaml", "style"
                                    fileName = Path.GetFileName(file)
                                    IO.File.Copy(file, Path.Combine(directory, fileName))
                            End Select
                        Next

                    Else
                        IO.File.Copy(libraryFile, Path.Combine(directory, fileName))
                        copySBLib = True
                    End If

                Catch
                End Try
            Next

            If copySBLib Then
                location = GetType(SmallBasic.Library.Primitive).Assembly.Location
                fileName = Path.GetFileName(location)

                Try
                    IO.File.Copy(location, Path.Combine(directory, fileName), overwrite:=False)
                Catch
                End Try
            End If
        End Sub

        Private Shared Sub PopulateReferences()
            _referenceAssemblies.Add(GetType(Primitive).Assembly)

            For Each reference In _references
                Try
                    Dim item = Assembly.LoadFile(reference)
                    _referenceAssemblies.Add(item)
                Catch
                    Throw New InvalidOperationException(String.Format(ResourceHelper.GetString("LoadReferenceFailed"), reference))
                End Try
            Next
        End Sub

        Private Shared Sub PopulateClrSymbols()
            For Each refAssembly In _referenceAssemblies
                AddAssemblyTypesToList(refAssembly)
            Next

            Dim exePath = Assembly.GetExecutingAssembly().Location
            Dim directoryName = IO.Path.GetDirectoryName(exePath)
            Dim libPath = IO.Path.Combine(directoryName, "lib")
            LoadDlls(libPath)

            Dim folderPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
            Dim appDataPath = IO.Path.Combine(folderPath, "Microsoft", "Small Basic", "Lib")
            LoadDlls(appDataPath)

        End Sub

        Private Shared Sub LoadExes(path As String)
            For Each dirName In Directory.GetDirectories(path)
                Dim files = Directory.GetFiles(dirName, "*.exe")

                For Each fileName In files
                    Try
                        Dim _assembly = Assembly.LoadFile(fileName)
                        Dim type = _assembly?.GetType("Global")
                        If type IsNot Nothing AndAlso type.IsVisible Then
                            Dim name = IO.Path.GetFileNameWithoutExtension(fileName)
                            name = name.Replace(" ", "")
                            AddTypeToList(type, name)
                            _libraryFiles.Add(fileName)
                        End If
                    Catch ex As Exception
                    End Try
                Next

            Next
        End Sub

        Private Shared Sub LoadDlls(path As String)
            If Not Directory.Exists(path) Then Return

            Dim files = Directory.GetFiles(path, "*.dll")
            For Each fileName In files
                Try
                    Dim _assembly = Assembly.LoadFile(fileName)
                    If AddAssemblyTypesToList(_assembly) Then
                        _libraryFiles.Add(fileName)
                    End If
                Catch ex As Exception
                End Try
            Next

            LoadExes(path)
        End Sub

        Private Shared Function AddAssemblyTypesToList(assembly As Assembly) As Boolean
            If assembly Is Nothing Then Return False

            Dim result = False
            Dim types = assembly.GetTypes()

            For Each type In types
                If type.IsVisible AndAlso IsSBType(type) Then
                    AddTypeToList(type)
                    result = True
                End If
            Next

            Return result
        End Function

        Private Shared Function IsSBType(type As Type) As Boolean
            Return type.GetCustomAttributes(GetType(SmallBasicTypeAttribute), inherit:=False).Length > 0 OrElse
                type.GetCustomAttributes(GetType(SmallBasic.Library.SmallBasicTypeAttribute), inherit:=False).Length > 0
        End Function

        Private Shared Sub AddTypeToList(type As Type, Optional typeName As String = Nothing)
            Dim typeInfo As New TypeInfo With {
                .Type = type,
                .Name = If(typeName, type.Name),
                .HideFromIntellisense = HideType(type)
            }

            Dim methods = type.GetMethods(BindingFlags.Static Or BindingFlags.Public)

            For Each methodInfo In methods
                If CanAddMethod(methodInfo) Then
                    Dim name = methodInfo.Name.ToLower()
                    If Not typeInfo.Methods.ContainsKey(name) Then
                        typeInfo.Methods.Add(name, methodInfo)
                    End If
                End If
            Next

            Dim properties = type.GetProperties(BindingFlags.Static Or BindingFlags.Public)

            For Each propertyInfo In properties
                If CanAddProperty(propertyInfo) Then
                    typeInfo.Properties.Add(propertyInfo.Name.ToLower(), propertyInfo)
                End If
            Next

            Dim events = type.GetEvents(BindingFlags.Static Or BindingFlags.Public)

            For Each eventInfo In events
                If CanAddEvent(eventInfo) Then
                    typeInfo.Events.Add(eventInfo.Name.ToLower(), eventInfo)
                End If
            Next

            If typeInfo.Events.Count > 0 OrElse typeInfo.Methods.Count > 0 OrElse typeInfo.Properties.Count > 0 Then
                _typeInfoBag.Types(typeInfo.Key) = typeInfo
            End If
        End Sub

        Private Shared Function HideType(type As Type) As Boolean
            Return type.GetCustomAttributes(GetType(HideFromIntellisenseAttribute), inherit:=False).Length > 0 OrElse
                type.GetCustomAttributes(GetType(SmallBasic.Library.HideFromIntellisenseAttribute), inherit:=False).Length > 0
        End Function

        Dim svbDir = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) & "\SmallVisualBasic"

        Private Shared Function CanAddMethod(methodInfo As MethodInfo) As Boolean
            If Not methodInfo.IsGenericMethod AndAlso
                    Not methodInfo.IsConstructor AndAlso
                    Not methodInfo.ContainsGenericParameters AndAlso
                    Not methodInfo.IsSpecialName AndAlso (
                        methodInfo.ReturnType Is GetType(Void) OrElse
                        IsPrimitive(methodInfo.ReturnType)
                    ) Then

                Dim parameters = methodInfo.GetParameters()

                For Each paramInfo In parameters
                    If Not IsPrimitive(paramInfo.ParameterType) Then
                        Return False
                    End If
                Next

                Return True
            End If

            Return False
        End Function

        Private Shared Function IsPrimitive(returnType As Type) As Boolean
            Return returnType Is GetType(Primitive) OrElse returnType Is GetType(SmallBasic.Library.Primitive)
        End Function

        Private Shared Function CanAddProperty(propertyInfo As PropertyInfo) As Boolean
            If Not propertyInfo.IsSpecialName Then
                Return IsPrimitive(propertyInfo.PropertyType)
            End If

            Return False
        End Function

        Private Shared Function CanAddEvent(eventInfo As EventInfo) As Boolean
            If eventInfo.IsSpecialName Then Return False

            Return eventInfo.EventHandlerType Is GetType(SmallBasicCallback) OrElse
                    eventInfo.EventHandlerType Is GetType(SmallBasic.Library.SmallBasicCallback)
        End Function

        Private Shared Sub PopulatePrimitiveMethods()
            Dim typeFromHandle = GetType(Primitive)
            _typeInfoBag.StringToPrimitive = typeFromHandle.GetMethod("op_Implicit", New Type(0) {GetType(String)})
            _typeInfoBag.NumberToPrimitive = typeFromHandle.GetMethod("op_Implicit", New Type(0) {GetType(Double)})
            _typeInfoBag.DateToPrimitive = typeFromHandle.GetMethod("DateToPrimitive")
            _typeInfoBag.TimeSpanToPrimitive = typeFromHandle.GetMethod("TimeSpanToPrimitive")
            _typeInfoBag.PrimitiveToBoolean = typeFromHandle.GetMethod("ConvertToBoolean")
            _typeInfoBag.Negation = typeFromHandle.GetMethod("op_UnaryNegation")
            _typeInfoBag.Add = typeFromHandle.GetMethod("op_Addition")
            _typeInfoBag.Subtract = typeFromHandle.GetMethod("op_Subtraction")
            _typeInfoBag.Multiply = typeFromHandle.GetMethod("op_Multiply")
            _typeInfoBag.Divide = typeFromHandle.GetMethod("op_Division")
            _typeInfoBag.GreaterThan = typeFromHandle.GetMethod("op_GreaterThan")
            _typeInfoBag.GreaterThanOrEqualTo = typeFromHandle.GetMethod("op_GreaterThanOrEqual")
            _typeInfoBag.LessThan = typeFromHandle.GetMethod("op_LessThan")
            _typeInfoBag.LessThanOrEqualTo = typeFromHandle.GetMethod("op_LessThanOrEqual")
            _typeInfoBag.EqualTo = typeFromHandle.GetMethod("op_Equality", New Type(1) {GetType(Primitive), GetType(Primitive)})
            _typeInfoBag.NotEqualTo = typeFromHandle.GetMethod("op_Inequality", New Type(1) {GetType(Primitive), GetType(Primitive)})
            _typeInfoBag.And = typeFromHandle.GetMethod("op_And", New Type(1) {GetType(Primitive), GetType(Primitive)})
            _typeInfoBag.Or = typeFromHandle.GetMethod("op_Or", New Type(1) {GetType(Primitive), GetType(Primitive)})
            _typeInfoBag.GetArrayValue = typeFromHandle.GetMethod("GetArrayValue")
            _typeInfoBag.SetArrayValue = typeFromHandle.GetMethod("SetArrayValue")
        End Sub
    End Class

End Namespace
