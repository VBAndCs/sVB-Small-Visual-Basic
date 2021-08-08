Imports System
Imports System.Collections.Generic
Imports System.Globalization
Imports System.IO
Imports System.Reflection
Imports Microsoft.SmallBasic.Library

Namespace Microsoft.SmallBasic
    Public Class Compiler
        Private _referenceAssemblies As List(Of Assembly)
        Private _libraryFiles As New List(Of String)()
        Private _errors As New List(Of [Error])()

        Public ReadOnly Property References As New List(Of String)
        Public ReadOnly Property Parser As Parser
        Public ReadOnly Property TypeInfoBag As TypeInfoBag

        Public Sub New()
            _TypeInfoBag = New TypeInfoBag()
            Initialize()
            _Parser = New Parser(_errors, _TypeInfoBag)
        End Sub

        Private Sub Initialize()
            PopulateReferences()
            PopulateClrSymbols()
            PopulatePrimitiveMethods()
        End Sub

        Public Function Compile(ByVal source As TextReader) As List(Of [Error])
            If source Is Nothing Then
                Throw New ArgumentNullException("source")
            End If

            _errors.Clear()
            _parser.Parse(source)

            Dim semanticAnalyzer As New SemanticAnalyzer(_parser, _typeInfoBag)
            semanticAnalyzer.Analyze()
            Return _errors
        End Function

        Public Function Build(ByVal source As TextReader, ByVal outputName As String, ByVal directory As String) As List(Of [Error])
            If source Is Nothing Then
                Throw New ArgumentNullException("source")
            End If

            If Equals(outputName, Nothing) Then
                Throw New ArgumentNullException("outputName")
            End If

            If Equals(directory, Nothing) Then
                Throw New ArgumentNullException("directory")
            End If

            Compile(source)

            If _errors.Count > 0 Then Return _errors

            Dim codeGenerator As New CodeGenerator(_parser, _typeInfoBag, outputName, directory)
            codeGenerator.GenerateExecutable()
            CopyLibraryAssemblies(directory)
            Return _errors
        End Function

        Private Sub CopyLibraryAssemblies(ByVal directory As String)
            Dim location = GetType(Primitive).Assembly.Location
            Dim fileName = Path.GetFileName(location)

            Try
                IO.File.Copy(location, Path.Combine(directory, fileName), overwrite:=True)
            Catch
            End Try

            For Each libraryFile In _libraryFiles

                Try
                    fileName = Path.GetFileName(libraryFile)
                    IO.File.Copy(libraryFile, Path.Combine(directory, fileName), overwrite:=True)
                Catch
                End Try
            Next
        End Sub

        Private Sub PopulateReferences()
            _referenceAssemblies = New List(Of Assembly)()
            _referenceAssemblies.Add(GetType(Primitive).Assembly)

            For Each reference In References
                Try
                    Dim item = Assembly.LoadFile(reference)
                    _referenceAssemblies.Add(item)
                Catch
                    Throw New InvalidOperationException(String.Format(ResourceHelper.GetString("LoadReferenceFailed"), reference))
                End Try
            Next
        End Sub

        Private Sub PopulateClrSymbols()
            For Each referenceAssembly In _referenceAssemblies
                AddAssemblyTypesToList(referenceAssembly)
            Next

            Dim directoryName As String = IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
            Dim path = IO.Path.Combine(directoryName, "lib")

            If Directory.Exists(path) Then
                Dim files = Directory.GetFiles(path, "*.dll")

                For Each fileName In files

                    Try
                        Dim _assembly = Assembly.LoadFile(fileName)
                        AddAssemblyTypesToList(_assembly)
                        _libraryFiles.Add(fileName)
                    Catch
                    End Try
                Next
            End If

            LoadAssembliesFromAppData()
        End Sub

        Private Sub LoadAssembliesFromAppData()
            Dim folderPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
            Dim path = IO.Path.Combine(folderPath, "Microsoft", "Small Basic", "Lib")

            If Not Directory.Exists(path) Then
                Return
            End If

            Dim files = Directory.GetFiles(path, "*.dll")

            For Each text In files

                Try
                    Dim _assembly = Assembly.LoadFile(text)
                    AddAssemblyTypesToList(_assembly)
                    _libraryFiles.Add(text)
                Catch
                End Try
            Next
        End Sub

        Private Function AddAssemblyTypesToList(ByVal assembly As Assembly) As Boolean
            If assembly Is Nothing Then Return False


            Dim result = False
            Dim types As Type() = assembly.GetTypes()

            For Each type In types

                If type.GetCustomAttributes(GetType(SmallBasicTypeAttribute), inherit:=False).Length > 0 AndAlso type.IsVisible Then
                    AddTypeToList(type)
                    result = True
                End If
            Next

            Return result
        End Function

        Private Sub AddTypeToList(ByVal type As Type)
            Dim typeInfo As New TypeInfo()
            typeInfo.Type = type
            typeInfo.HideFromIntellisense = type.GetCustomAttributes(GetType(HideFromIntellisenseAttribute), inherit:=False).Length > 0
            Dim methods = type.GetMethods(BindingFlags.Static Or BindingFlags.Public)

            For Each methodInfo In methods

                If CanAddMethod(methodInfo) AndAlso Not typeInfo.Methods.ContainsKey(methodInfo.Name.ToLower(CultureInfo.CurrentUICulture)) Then
                    typeInfo.Methods.Add(methodInfo.Name.ToLower(CultureInfo.CurrentUICulture), methodInfo)
                End If
            Next

            'New Dictionary(Of String, PropertyInfo)()
            Dim properties = type.GetProperties(BindingFlags.Static Or BindingFlags.Public)

            For Each propertyInfo In properties

                If CanAddProperty(propertyInfo) Then
                    typeInfo.Properties.Add(propertyInfo.Name.ToLower(CultureInfo.CurrentUICulture), propertyInfo)
                End If
            Next

            'New Dictionary(Of String, EventInfo)()
            Dim events = type.GetEvents(BindingFlags.Static Or BindingFlags.Public)

            For Each eventInfo In events

                If CanAddEvent(eventInfo) Then
                    typeInfo.Events.Add(eventInfo.Name.ToLower(CultureInfo.CurrentUICulture), eventInfo)
                End If
            Next

            If typeInfo.Events.Count > 0 OrElse typeInfo.Methods.Count > 0 OrElse typeInfo.Properties.Count > 0 Then
                _typeInfoBag.Types(type.Name.ToLower(CultureInfo.CurrentUICulture)) = typeInfo
            End If
        End Sub

        Private Function CanAddMethod(ByVal methodInfo As MethodInfo) As Boolean
            If Not methodInfo.IsGenericMethod AndAlso
                    Not methodInfo.IsConstructor AndAlso
                    Not methodInfo.ContainsGenericParameters AndAlso
                    Not methodInfo.IsSpecialName AndAlso
                    (methodInfo.ReturnType Is GetType(Void) OrElse
                     methodInfo.ReturnType Is GetType(Primitive)) Then

                Dim parameters = methodInfo.GetParameters()

                For Each paramInfo In parameters
                    If paramInfo.ParameterType IsNot GetType(Primitive) Then Return False
                Next

                Return True
            End If

            Return False
        End Function

        Private Function CanAddProperty(ByVal propertyInfo As PropertyInfo) As Boolean
            If Not propertyInfo.IsSpecialName Then
                Return propertyInfo.PropertyType Is GetType(Primitive)
            End If

            Return False
        End Function

        Private Function CanAddEvent(ByVal eventInfo As EventInfo) As Boolean
            If Not eventInfo.IsSpecialName Then
                Return eventInfo.EventHandlerType Is GetType(SmallBasicCallback)
            End If

            Return False
        End Function

        Private Sub PopulatePrimitiveMethods()
            Dim typeFromHandle = GetType(Primitive)
            _typeInfoBag.StringToPrimitive = typeFromHandle.GetMethod("op_Implicit", New Type(0) {GetType(String)})
            _typeInfoBag.NumberToPrimitive = typeFromHandle.GetMethod("op_Implicit", New Type(0) {GetType(Double)})
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
