Imports System.ComponentModel.Composition

Namespace Microsoft.Nautilus.Core
    <ExportableAttribute>
    <AttributeUsage(AttributeTargets.Class Or AttributeTargets.Method Or AttributeTargets.Property Or AttributeTargets.Field, AllowMultiple:=True)>
    Public Class BaseMetadataProviderAttribute
        Inherits Attribute
    End Class
End Namespace
