Imports System
Imports System.Reflection
Imports System.Reflection.Emit
Imports Microsoft.SmallVisualBasic.Library
Imports Microsoft.SmallVisualBasic.Library.Internal

Namespace Microsoft.SmallVisualBasic
    Public Class CodeGenerator
        Private _outputName As String
        Private _directory As String
        Private _parsers As List(Of Parser)
        Private _entryPoint As MethodInfo
        Private _currentScope As CodeGenScope
        Private _typeInfoBag As TypeInfoBag

        Public Sub New(parsers As List(Of Parser), typeInfoBag As TypeInfoBag, outputName As String, directory As String)

            If typeInfoBag Is Nothing Then Throw New ArgumentNullException("typeInfoBag")

            _parsers = parsers
            _typeInfoBag = typeInfoBag
            _outputName = outputName
            _directory = directory
        End Sub

        Public Shared IgnoreVarErrors As Boolean

        Public Shared Sub LowerAndEmit(code As String, scope As CodeGenScope, Subroutine As Statements.SubroutineStatement, lineOffset As Integer)
            IgnoreVarErrors = True
            Dim tempRoutine = Statements.SubroutineStatement.Current
            Statements.SubroutineStatement.Current = Subroutine
            Dim _parser = Parser.Parse(code, scope.SymbolTable, scope.TypeInfoBag, lineOffset)

            For Each item In _parser.ParseTree
                item.PrepareForEmit(scope)
            Next


            For Each item In _parser.ParseTree
                item.EmitIL(scope)
            Next

            IgnoreVarErrors = False
            Statements.SubroutineStatement.Current = tempRoutine

        End Sub

        Public Function GenerateExecutable() As Boolean
            Dim assemblyName As New AssemblyName()
            assemblyName.Name = _outputName
            Dim assemblyBuilder = AppDomain.CurrentDomain.DefineDynamicAssembly(assemblyName, AssemblyBuilderAccess.Save, _directory)
            Dim moduleBuilder = assemblyBuilder.DefineDynamicModule(_outputName & ".exe", emitSymbolInfo:=True)

            Dim mainFormInitialize As MethodInfo
            Dim formInit As MethodInfo
            For Each parser In _parsers
                formInit = EmitModule(parser, moduleBuilder)
                If parser.IsMainForm Then mainFormInitialize = formInit
            Next

            If Not EmiMain(If(mainFormInitialize, formInit), moduleBuilder) Then Return False

            assemblyBuilder.SetEntryPoint(_entryPoint, PEFileKinds.WindowApplication)
            assemblyBuilder.Save(_outputName & ".exe")
            Return True
        End Function

        Private Function EmitModule(parser As Parser, moduleBuilder As ModuleBuilder) As MethodInfo
            Dim typeBuilder = moduleBuilder.DefineType(parser.ClassName, TypeAttributes.Sealed)
            Dim methodBuilder = typeBuilder.DefineMethod("Initialize", MethodAttributes.Static Or MethodAttributes.Public)
            Dim iLGenerator = methodBuilder.GetILGenerator()
            _currentScope = New CodeGenScope With {
                .ILGenerator = iLGenerator,
                .MethodBuilder = methodBuilder,
                .TypeBuilder = typeBuilder,
                .SymbolTable = parser.SymbolTable,
                .TypeInfoBag = _typeInfoBag
            }

            BuildFields(typeBuilder, parser.SymbolTable)
            EmitIL(parser.ParseTree)
            iLGenerator.Emit(OpCodes.Ret)
            typeBuilder.CreateType()
            Return methodBuilder
        End Function

        Private Function EmiMain(formInit As MethodInfo, moduleBuilder As ModuleBuilder) As Boolean
            Dim typeBuilder = moduleBuilder.DefineType("_SmallVisualBasic_Program", TypeAttributes.Sealed)
            _entryPoint = typeBuilder.DefineMethod("_Main", MethodAttributes.Static)
            Dim methodBuilder = CType(_entryPoint, MethodBuilder)
            Dim iLGenerator = methodBuilder.GetILGenerator()
            iLGenerator.EmitCall(OpCodes.Call, GetType(SmallBasicApplication).GetMethod("BeginProgram", BindingFlags.Static Or BindingFlags.Public), Nothing)
            iLGenerator.EmitCall(OpCodes.Call, formInit, Nothing)
            iLGenerator.EmitCall(OpCodes.Call, GetType(TextWindow).GetMethod("PauseIfVisible", BindingFlags.Static Or BindingFlags.Public), Nothing)
            iLGenerator.Emit(OpCodes.Ret)
            typeBuilder.CreateType()
            Return True
        End Function

        Private Sub BuildFields(typeBuilder As TypeBuilder, symbolTable As SymbolTable)
            For Each key In symbolTable.GlobalVariables.Keys
                Dim value = typeBuilder.DefineField(key, GetType(Primitive), FieldAttributes.Private Or FieldAttributes.Static)
                _currentScope.Fields.Add(key, value)
            Next
        End Sub

        Private Sub EmitIL(parseTree As List(Of Statements.Statement))
            For Each item In parseTree
                item.PrepareForEmit(_currentScope)
            Next

            For Each item In parseTree
                item.EmitIL(_currentScope)
            Next
        End Sub

    End Class
End Namespace
