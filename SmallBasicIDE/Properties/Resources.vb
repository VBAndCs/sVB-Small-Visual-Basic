Imports System.CodeDom.Compiler
Imports System.ComponentModel
Imports System.Diagnostics
Imports System.Globalization
Imports System.Resources
Imports System.Runtime.CompilerServices

Namespace Microsoft.SmallBasic.Properties
    <CompilerGenerated>
    <DebuggerNonUserCode>
    <GeneratedCode("System.Resources.Tools.StronglyTypedResourceBuilder", "4.0.0.0")>
    Friend Class Resources
        Private Shared resourceMan As ResourceManager
        Private Shared resourceCulture As CultureInfo

        <EditorBrowsable(EditorBrowsableState.Advanced)>
        Friend Shared ReadOnly Property ResourceManager As ResourceManager
            Get

                If resourceMan Is Nothing Then
                    resourceMan = New ResourceManager("Microsoft.SmallBasic.Properties.Resources", GetType(Resources).Assembly)
                End If

                Return resourceMan
            End Get
        End Property

        <EditorBrowsable(EditorBrowsableState.Advanced)>
        Friend Shared Property Culture As CultureInfo
            Get
                Return resourceCulture
            End Get
            Set(value As CultureInfo)
                resourceCulture = value
            End Set
        End Property

        Friend Sub New()
        End Sub

    End Class
End Namespace
