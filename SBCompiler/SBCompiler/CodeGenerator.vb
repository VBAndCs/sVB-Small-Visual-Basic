Imports System
Imports System.Reflection
Imports System.Reflection.Emit
Imports Microsoft.SmallBasic.Library
Imports Microsoft.SmallBasic.Library.Internal

Namespace Microsoft.SmallBasic
    Public Class CodeGenerator
        Private _outputName As String
        Private _directory As String
        Private _parser As Parser
        Private _entryPoint As MethodInfo
        Private _currentScope As CodeGenScope
        Private _typeInfoBag As TypeInfoBag
        Private _symbolTable As SymbolTable

        Public Sub New(parser As Parser, typeInfoBag As TypeInfoBag, outputName As String, directory As String)
            If parser Is Nothing Then Throw New ArgumentNullException("parser")

            If typeInfoBag Is Nothing Then Throw New ArgumentNullException("typeInfoBag")

            _parser = parser
            _symbolTable = _parser.SymbolTable
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

            'Dim semantic As New SemanticAnalyzer(_parser, scope.TypeInfoBag)
            'semantic.Analyze()

            ' EmitIL
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

            If Not EmitModule(moduleBuilder) Then Return False

            assemblyBuilder.SetEntryPoint(_entryPoint, PEFileKinds.WindowApplication)
            assemblyBuilder.Save(_outputName & ".exe")
            Return True
        End Function

        Private Function EmitModule(moduleBuilder As ModuleBuilder) As Boolean
            Dim typeBuilder = moduleBuilder.DefineType("_SmallBasicProgram", TypeAttributes.Sealed)
            _entryPoint = typeBuilder.DefineMethod("_Main", MethodAttributes.Static)
            Dim methodBuilder = CType(_entryPoint, MethodBuilder)
            Dim iLGenerator = methodBuilder.GetILGenerator()
            _currentScope = New CodeGenScope With {
                .ILGenerator = iLGenerator,
                .MethodBuilder = methodBuilder,
                .TypeBuilder = typeBuilder,
                .SymbolTable = _symbolTable,
                .TypeInfoBag = _typeInfoBag
            }

            BuildFields(typeBuilder)
            iLGenerator.EmitCall(OpCodes.Call, GetType(SmallBasicApplication).GetMethod("BeginProgram", BindingFlags.Static Or BindingFlags.Public), Nothing)
            EmitIL()
            iLGenerator.EmitCall(OpCodes.Call, GetType(TextWindow).GetMethod("PauseIfVisible", BindingFlags.Static Or BindingFlags.Public), Nothing)

            ' Bad for win forms
            'iLGenerator.EmitCall(OpCodes.Call, GetType(SmallBasicApplication).GetMethod("EndProgram", BindingFlags.Static Or BindingFlags.Public), Nothing)

            iLGenerator.Emit(OpCodes.Ret)
            typeBuilder.CreateType()
            Return True
        End Function

        Private Sub BuildFields(typeBuilder As TypeBuilder)
            Dim symbolTable = _parser.SymbolTable

            For Each key In symbolTable.GlobalVariables.Keys
                Dim value = typeBuilder.DefineField(key, GetType(Primitive), FieldAttributes.Private Or FieldAttributes.Static)
                _currentScope.Fields.Add(key, value)
            Next
        End Sub

        Private Sub EmitIL()
            For Each item In _parser.ParseTree
                item.PrepareForEmit(_currentScope)
            Next

            For Each item In _parser.ParseTree
                item.EmitIL(_currentScope)
            Next
        End Sub

    End Class
End Namespace
