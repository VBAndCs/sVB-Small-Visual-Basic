Imports Microsoft.SmallVisualBasic.Library

Namespace WinForms

    ''' <summary>
    ''' Contains the names of common dialog results
    ''' </summary>
    <SmallBasicType>
    Public NotInheritable Class DialogResults
        <ReturnValueType(VariableType.DialogResult)>
        Public Shared ReadOnly Property Cancel As Primitive = "Cancel"

        <ReturnValueType(VariableType.DialogResult)>
        Public Shared ReadOnly Property OK As Primitive = "OK"

        <ReturnValueType(VariableType.DialogResult)>
        Public Shared ReadOnly Property Yes As Primitive = "Yes"

        <ReturnValueType(VariableType.DialogResult)>
        Public Shared ReadOnly Property No As Primitive = "No"

        <ReturnValueType(VariableType.DialogResult)>
        Public Shared ReadOnly Property Abort As Primitive = "Abort"

        <ReturnValueType(VariableType.DialogResult)>
        Public Shared ReadOnly Property Retry As Primitive = "Retry"

        <ReturnValueType(VariableType.DialogResult)>
        Public Shared ReadOnly Property Ignore As Primitive = "Ignore"

        <ReturnValueType(VariableType.DialogResult)>
        Public Shared ReadOnly Property [Try] As Primitive = "Try"

        <ReturnValueType(VariableType.DialogResult)>
        Public Shared ReadOnly Property [Continue] As Primitive = "Continue"

    End Class

End Namespace
