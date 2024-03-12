Imports System.Reflection
Imports System.Reflection.Emit
Imports Microsoft.SmallVisualBasic.Library
Imports Microsoft.SmallVisualBasic.Library.Internal

Namespace Microsoft.SmallVisualBasic
    Public Class CodeGenerator
        Private _outputName As String
        Private _directory As String
        Private _appName As String
        Private _parsers As List(Of Parser)
        Private Shared entryPoint As MethodInfo
        Private _currentScope As CodeGenScope
        Private _typeInfoBag As TypeInfoBag

        Public Sub New(
                          parsers As List(Of Parser),
                          typeInfoBag As TypeInfoBag,
                          outputName As String,
                          directory As String
                    )

            If parsers Is Nothing Then
                Throw New ArgumentNullException(NameOf(parsers))
            End If

            If typeInfoBag Is Nothing Then
                Throw New ArgumentNullException(NameOf(typeInfoBag))
            End If

            _parsers = parsers
            _typeInfoBag = typeInfoBag
            _outputName = Char.ToUpper(outputName(0)) & If(outputName.Length > 1, outputName.Substring(1), "")
            _appName = "Global"
            _directory = directory

            _xmlDoc = New XDocument
            Dim doc =
                <doc>
                    <assembly>
                        <name><%= _appName %></name>
                    </assembly>
                </doc>

            _members = <members/>
                doc.Add(_members)
                _xmlDoc.Add(doc)

        End Sub

        Public Shared IgnoreVarErrors As Boolean

        Public Shared Sub LowerAndEmit(code As String, scope As CodeGenScope, Subroutine As Statements.SubroutineStatement, lineOffset As Integer)
            IgnoreVarErrors = True
            Dim tempRoutine = Statements.SubroutineStatement.Current
            Statements.SubroutineStatement.Current = Subroutine
            Dim _parser = Parser.Parse(code, scope.SymbolTable, lineOffset)

            For Each item In _parser.ParseTree
                item.PrepareForEmit(scope)
            Next


            For Each item In _parser.ParseTree
                item.EmitIL(scope)
            Next

            IgnoreVarErrors = False
            Statements.SubroutineStatement.Current = tempRoutine

        End Sub

        Private Sub AddGlobalTypeToList(type As Type)
            Dim typeInfo As New TypeInfo
            typeInfo.Type = type

            Dim methods = type.GetMethods(BindingFlags.Public Or BindingFlags.Static)
            For Each method In methods
                If Not method.IsSpecialName Then
                    Dim name = method.Name.ToLower()
                    If name <> "initialize" Then
                        typeInfo.Methods.Add(name, method)
                    End If
                End If
            Next

            Dim props = type.GetProperties(BindingFlags.Public Or BindingFlags.Static)
            For Each prop In props
                Dim name = prop.Name.ToLower()
                typeInfo.Properties.Add(name, prop)
            Next

            If typeInfo.Methods.Count > 0 OrElse typeInfo.Properties.Count > 0 Then
                _typeInfoBag.Types("global") = typeInfo
            End If
        End Sub

        Public Function GenerateExecutable(Optional forGlobalHelp As Boolean = False) As AssemblyName()
            Dim asmName As New AssemblyName(_outputName)
            Dim asm = AppDomain.CurrentDomain.DefineDynamicAssembly(
                asmName,
                AssemblyBuilderAccess.Save,
                _directory
            )

            Dim moduleBuilder = asm.DefineDynamicModule(_outputName & ".exe", emitSymbolInfo:=True)
            Dim typeBuilder = moduleBuilder.DefineType("_SmallVisualBasic_Program", TypeAttributes.Sealed)
            EntryPoint = typeBuilder.DefineMethod("_Main", MethodAttributes.Static)

            Dim mainFormInit As MethodInfo = Nothing
            Dim formInit As MethodInfo = Nothing

            For Each parser In _parsers
                formInit = EmitModule(parser, moduleBuilder, forGlobalHelp)
                If parser.IsMainForm Then
                    mainFormInit = formInit
                ElseIf parser.IsGlobal Then
                    AddGlobalTypeToList(formInit.DeclaringType)
                    AddMethodDocs(parser)
                End If
            Next

            If Not forGlobalHelp Then
                EmitMain(If(mainFormInit, formInit))
                typeBuilder.CreateType()
                asm.SetEntryPoint(EntryPoint, PEFileKinds.WindowApplication)
                asm.Save(_outputName & ".exe")
                _xmlDoc.Save(IO.Path.Combine(_directory, _outputName & ".xml"))
                Return asm.GetReferencedAssemblies()
            End If

            Return Nothing
        End Function



        Dim globalScope As CodeGenScope

        Private Function EmitModule(
                          parser As Parser,
                          moduleBuilder As ModuleBuilder,
                          Optional forGlobalHelp As Boolean = False
                    ) As MethodInfo

            Dim domain = TypeAttributes.Sealed
            If parser.IsGlobal Then
                domain = domain Or TypeAttributes.Public
            End If

            Dim typeBuilder = moduleBuilder.DefineType(parser.ClassName, domain)
            Dim initializeMethod = typeBuilder.DefineMethod(
                "Initialize",
                MethodAttributes.Static Or MethodAttributes.Public
            )
            Dim attr = GetType(HideFromIntellisenseAttribute)
            Dim ctorInfo = attr.GetConstructor(Type.EmptyTypes)
            initializeMethod.SetCustomAttribute(
                 New CustomAttributeBuilder(ctorInfo, New Object() {})
            )

            Dim iL = initializeMethod.GetILGenerator()
            _currentScope = New CodeGenScope With {
                .ILGenerator = iL,
                .MethodBuilder = initializeMethod,
                .TypeBuilder = typeBuilder,
                .SymbolTable = parser.SymbolTable,
                .TypeInfoBag = _typeInfoBag,
                .GlobalScope = globalScope,
                .ForGlobalHelp = forGlobalHelp
            }

            If parser.IsGlobal Then
                globalScope = _currentScope
                _currentScope.GlobalScope = globalScope
                If Not forGlobalHelp Then AddTypeDoc(parser)
            End If

            BuildFields(typeBuilder, parser.SymbolTable, parser.IsGlobal)
            EmitIL(parser.ParseTree)
            iL.Emit(OpCodes.Ret)

            If Not forGlobalHelp AndAlso parser.IsGlobal Then
                ' Create a shared constructor that calls the Inialize method
                Dim constructor = typeBuilder.DefineTypeInitializer()
                iL = constructor.GetILGenerator()
                iL.EmitCall(OpCodes.Call, initializeMethod, Nothing)
                iL.Emit(OpCodes.Ret)
            End If

            typeBuilder.CreateType()
            Return initializeMethod
        End Function

        Private Function EmitMain(mainFormInit As MethodInfo) As Boolean
            Dim methodBuilder = CType(entryPoint, MethodBuilder)
            Dim iLGenerator = methodBuilder.GetILGenerator()
            Dim beginProgram = GetType(SmallBasicApplication).GetMethod(
                "BeginProgram",
                BindingFlags.Static Or BindingFlags.Public
            )
            iLGenerator.EmitCall(OpCodes.Call, beginProgram, Nothing)
            iLGenerator.EmitCall(OpCodes.Call, mainFormInit, Nothing)

            Dim pauseThenClose = GetType(TextWindow).GetMethod(
               NameOf(TextWindow.PauseThenClose),
                BindingFlags.Static Or BindingFlags.Public
             )
            iLGenerator.EmitCall(OpCodes.Call, pauseThenClose, Nothing)

            iLGenerator.Emit(OpCodes.Ret)
            Return True
        End Function

        Private Sub BuildFields(
                        typeBuilder As TypeBuilder,
                        symbolTable As SymbolTable,
                        isGlobal As Boolean
                    )

            For Each var In symbolTable.GlobalVariables
                Dim fieldName = var.Value.LCaseText
                Dim fieldBuilder = typeBuilder.DefineField(
                        "_p_" & fieldName,
                        GetType(Primitive),
                        FieldAttributes.Private Or FieldAttributes.Static
                )
                _currentScope.Fields.Add(var.Key, fieldBuilder)

                ' private fields starts with _. They will be hidden
                If fieldName.StartsWith("_") Then Continue For

                If isGlobal Then
                    If Not _currentScope.ForGlobalHelp Then
                        AddPropertyDoc(var.Value)
                    End If

                    ' Define a public property for the field
                    Dim propName = var.Value.Text
                    Dim propBuilder = typeBuilder.DefineProperty(
                            propName,
                            PropertyAttributes.None,
                            GetType(Primitive),
                            Nothing
                    )

                    Dim attr = MethodAttributes.Public Or
                                      MethodAttributes.Static Or
                                      MethodAttributes.SpecialName Or
                                      MethodAttributes.HideBySig

                    Dim getProp = typeBuilder.DefineMethod(
                            $"get_{propName}",
                            attr,
                            GetType(Primitive),
                            Type.EmptyTypes
                    )

                    Dim getIL = getProp.GetILGenerator()
                    getIL.Emit(OpCodes.Ldsfld, fieldBuilder)
                    getIL.Emit(OpCodes.Ret)

                    Dim setProp = typeBuilder.DefineMethod(
                            $"set_{propName}",
                            attr,
                            Nothing,
                            New Type() {GetType(Primitive)}
                    )

                    Dim setIL = setProp.GetILGenerator()
                    setIL.Emit(OpCodes.Ldarg_0)
                    setIL.Emit(OpCodes.Stsfld, fieldBuilder)
                    setIL.Emit(OpCodes.Ret)

                    propBuilder.SetGetMethod(getProp)
                    propBuilder.SetSetMethod(setProp)

                    Dim returntype = _currentScope.SymbolTable.GetInferedType(var.Value)
                    If returntype <> VariableType.Any Then
                        Dim ctorParams = New Type() {GetType(VariableType)}
                        Dim ctorInfo = GetType(WinForms.ReturnValueTypeAttribute).GetConstructor(ctorParams)
                        propBuilder.SetCustomAttribute(
                            New CustomAttributeBuilder(
                                    ctorInfo,
                                    New Object() {returntype}
                            )
                        )
                    End If

                End If
            Next
        End Sub

        Dim _xmlDoc As XDocument
        Dim _members As XElement

        Private Sub AddTypeDoc(parser As Parser)
            Dim atStart = True
            Dim sb As New System.Text.StringBuilder()
            For Each st In parser.ParseTree
                If TypeOf st Is Statements.EmptyStatement Then
                    Dim comment = st.EndingComment.Text
                    If comment = "" Then
                        If Not atStart Then Exit For
                    Else
                        atStart = False
                        sb.AppendLine(comment)
                    End If
                Else
                    Exit For
                End If
            Next

            Dim typeInfo = sb.ToString().Trim()
            If typeInfo <> "" Then
                Dim typeDoc =
                    <member name=<%= "T:" & _appName %>>
                        <summary><%= typeInfo %></summary>
                    </member>
                _members.Add(typeDoc)
            End If
        End Sub

        Private Sub AddPropertyDoc(token As Token)
            If token.Comment = "" Then Return

            Dim propDoc =
                <member name=<%= "P:" & _appName & "." & token.Text %>>
                    <summary><%= token.Comment %></summary>
                </member>
            _members.Add(propDoc)
        End Sub

        Private Sub AddMethodDocs(parser As Parser)
            Dim prmtv = GetType(Primitive).FullName
            For Each st In parser.ParseTree
                Dim method = TryCast(st, Statements.SubroutineStatement)
                If method Is Nothing Then Continue For

                Dim hasDoc As Boolean = False
                Dim paramsCount = If(method.Params Is Nothing, 0, method.Params.Count)
                Dim paramsList = If(
                    paramsCount = 0,
                    "",
                    "(" & String.Join(",", Enumerable.Repeat(prmtv, paramsCount)) & ")"
                )

                Dim methodDoc =
                    <member name=<%= $"M:{_appName}.{method.Name.Text}{paramsList}" %>/>

                Dim summery = method.GetSummery()
                If summery <> "" Then
                    hasDoc = True
                    methodDoc.Add(
                          <summary><%= summery %></summary>
                    )
                End If

                If paramsCount > 0 Then
                    For Each param In method.Params
                        Dim paramInfo = param.Comment
                        If paramInfo <> "" Then
                            hasDoc = True
                            methodDoc.Add(
                                <param name=<%= param.Text %>><%= paramInfo %></param>
                            )
                        End If
                    Next
                End If

                If method.SubToken.Type = TokenType.Function Then
                    Dim returnInfo = method.GetRetunDoc()
                    If returnInfo <> "" Then
                        hasDoc = True
                        methodDoc.Add(
                            <returns><%= returnInfo %></returns>
                        )
                    End If
                End If

                If hasDoc Then _members.Add(methodDoc)
            Next
        End Sub

        Private Sub EmitIL(parseTree As List(Of Statements.Statement), Optional prepareOnly As Boolean = False)
            For Each item In parseTree
                item.PrepareForEmit(_currentScope)
            Next

            If prepareOnly Then Return

            For Each item In parseTree
                item.EmitIL(_currentScope)
            Next
        End Sub

    End Class
End Namespace
