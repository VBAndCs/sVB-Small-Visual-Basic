Imports System.CodeDom.Compiler
Imports System.Configuration
Imports System.Diagnostics
Imports System.Runtime.CompilerServices

Namespace Microsoft.SmallVisualBasic.Properties
    <GeneratedCode("Microsoft.VisualStudio.Editors.SettingsDesigner.SettingsSingleFileGenerator", "10.0.0.0")>
    <CompilerGenerated>
    Friend NotInheritable Class Settings
        Inherits ApplicationSettingsBase

        Private Shared defaultInstance As Settings = CType(Synchronized(New Settings()), Settings)

        Public Shared ReadOnly Property [Default] As Settings
            Get
                Return defaultInstance
            End Get
        End Property

        <DefaultSettingValue("http://smallbasic.com/smallbasic.com/program/service.asmx")>
        <SpecialSetting(SpecialSetting.WebServiceUrl)>
        <DebuggerNonUserCode>
        <ApplicationScopedSetting>
        Public ReadOnly Property SB_com_smallbasic_Service As String
            Get
                Return CStr(Me("SB_com_smallbasic_Service"))
            End Get
        End Property
    End Class
End Namespace
