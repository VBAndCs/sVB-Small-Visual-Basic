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

        Public Sub New(ByVal parser As Parser, ByVal typeInfoBag As TypeInfoBag, ByVal outputName As String, ByVal directory As String)
            If parser Is Nothing Then
                Throw New ArgumentNullException("parser")
            End If

            If typeInfoBag Is Nothing Then
                Throw New ArgumentNullException("typeInfoBag")
            End If

            _parser = parser
            _symbolTable = _parser.SymbolTable
            _typeInfoBag = typeInfoBag
            _outputName = outputName
            _directory = directory
        End Sub

        Public Function GenerateExecutable() As Boolean
            Dim assemblyName As AssemblyName = New AssemblyName()
            assemblyName.Name = _outputName
            Dim assemblyBuilder = AppDomain.CurrentDomain.DefineDynamicAssembly(assemblyName, AssemblyBuilderAccess.Save, _directory)
            Dim moduleBuilder = assemblyBuilder.DefineDynamicModule(_outputName & ".exe", emitSymbolInfo:=True)

            If Not EmitModule(moduleBuilder) Then
                Return False
            End If

            assemblyBuilder.SetEntryPoint(_entryPoint, PEFileKinds.WindowApplication)
            assemblyBuilder.Save(_outputName & ".exe")
            Return True
        End Function

        Private Function EmitModule(ByVal moduleBuilder As ModuleBuilder) As Boolean
            Dim typeBuilder = moduleBuilder.DefineType("_SmallBasicProgram", TypeAttributes.Sealed)
            Dim methodBuilder = CType(CSharpImpl.__Assign(_entryPoint, typeBuilder.DefineMethod("_Main", MethodAttributes.Static)), MethodBuilder)
            Dim iLGenerator As ILGenerator = methodBuilder.GetILGenerator()
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
            iLGenerator.EmitCall(OpCodes.Call, GetType(SmallBasicApplication).GetMethod("EndProgram", BindingFlags.Static Or BindingFlags.Public), Nothing)
            iLGenerator.Emit(OpCodes.Ret)
            typeBuilder.CreateType()
            Return True
        End Function

        Private Sub BuildFields(ByVal typeBuilder As TypeBuilder)
            Dim symbolTable = _parser.SymbolTable

            For Each key In symbolTable.Variables.Keys
                Dim value As FieldInfo = typeBuilder.DefineField(key, GetType(Primitive), FieldAttributes.Private Or FieldAttributes.Static)
                _currentScope.Fields.Add(key, value)
            Next
        End Sub

        Private Sub EmitIL()
            For Each item In _parser.ParseTree
                item.PrepareForEmit(_currentScope)
            Next

            For Each item2 In _parser.ParseTree
                item2.EmitIL(_currentScope)
            Next
        End Sub

        Private Class CSharpImpl
            <Obsolete("Please refactor calling code to use normal Visual Basic assignment")>
            Shared Function __Assign(Of T)(ByRef target As T, value As T) As T
                target = value
                Return value
            End Function
        End Class
    End Class
End Namespace
