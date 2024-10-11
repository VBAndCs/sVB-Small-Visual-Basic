Imports Microsoft.SmallVisualBasic.Library

Namespace WinForms

    ''' <summary>
    ''' Contains the names of common dialog results
    ''' </summary>
    <SmallVisualBasicType>
    Public NotInheritable Class DialogResults

        ''' <summary>
        ''' The user clicked the Cancel button or the form close button
        ''' </summary>
        <ReturnValueType(VariableType.DialogResult)>
        Public Shared ReadOnly Property Cancel As New Primitive("Cancel")

        ''' <summary>
        ''' The user clicked the OK button.
        ''' </summary>
        <ReturnValueType(VariableType.DialogResult)>
        Public Shared ReadOnly Property OK As New Primitive("OK")

        ''' <summary>
        ''' The user clicked the Yes button.
        ''' </summary>
        <ReturnValueType(VariableType.DialogResult)>
        Public Shared ReadOnly Property Yes As New Primitive("Yes")

        ''' <summary>
        ''' The user clicked the No button.
        ''' </summary>
        <ReturnValueType(VariableType.DialogResult)>
        Public Shared ReadOnly Property No As New Primitive("No")

        ''' <summary>
        ''' The user clicked the Abort button.
        ''' </summary>
        <ReturnValueType(VariableType.DialogResult)>
        Public Shared ReadOnly Property Abort As New Primitive("Abort")

        ''' <summary>
        ''' The user clicked the Retry button.
        ''' </summary>
        <ReturnValueType(VariableType.DialogResult)>
        Public Shared ReadOnly Property Retry As New Primitive("Retry")

        ''' <summary>
        ''' The user clicked the Ignore button.
        ''' </summary>
        <ReturnValueType(VariableType.DialogResult)>
        Public Shared ReadOnly Property Ignore As New Primitive("Ignore")

        ''' <summary>
        ''' The user clicked the Try button.
        ''' </summary>
        <ReturnValueType(VariableType.DialogResult)>
        Public Shared ReadOnly Property [Try] As New Primitive("Try")

        ''' <summary>
        ''' The user clicked the Continue button.
        ''' </summary>
        <ReturnValueType(VariableType.DialogResult)>
        Public Shared ReadOnly Property [Continue] As New Primitive("Continue")

    End Class

End Namespace
