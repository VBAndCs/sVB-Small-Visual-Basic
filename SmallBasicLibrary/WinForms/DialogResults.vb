Imports Microsoft.SmallVisualBasic.Library

Namespace WinForms

    ''' <summary>
    ''' Contains the names of common dialog results
    ''' </summary>
    <SmallBasicType>
    Public NotInheritable Class DialogResults

        ''' <summary>
        ''' The use clicked the Cancel button or the form close button
        ''' </summary>
        <ReturnValueType(VariableType.DialogResult)>
        Public Shared ReadOnly Property Cancel As Primitive = "Cancel"

        ''' <summary>
        ''' The use clicked the OK button.
        ''' </summary>
        <ReturnValueType(VariableType.DialogResult)>
        Public Shared ReadOnly Property OK As Primitive = "OK"

        ''' <summary>
        ''' The use clicked the Yes button.
        ''' </summary>
        <ReturnValueType(VariableType.DialogResult)>
        Public Shared ReadOnly Property Yes As Primitive = "Yes"

        ''' <summary>
        ''' The use clicked the No button.
        ''' </summary>
        <ReturnValueType(VariableType.DialogResult)>
        Public Shared ReadOnly Property No As Primitive = "No"

        ''' <summary>
        ''' The use clicked the Abort button.
        ''' </summary>
        <ReturnValueType(VariableType.DialogResult)>
        Public Shared ReadOnly Property Abort As Primitive = "Abort"

        ''' <summary>
        ''' The use clicked the Retry button.
        ''' </summary>
        <ReturnValueType(VariableType.DialogResult)>
        Public Shared ReadOnly Property Retry As Primitive = "Retry"

        ''' <summary>
        ''' The use clicked the Ignore button.
        ''' </summary>
        <ReturnValueType(VariableType.DialogResult)>
        Public Shared ReadOnly Property Ignore As Primitive = "Ignore"

        ''' <summary>
        ''' The use clicked the Try button.
        ''' </summary>
        <ReturnValueType(VariableType.DialogResult)>
        Public Shared ReadOnly Property [Try] As Primitive = "Try"

        ''' <summary>
        ''' The use clicked the Continue button.
        ''' </summary>
        <ReturnValueType(VariableType.DialogResult)>
        Public Shared ReadOnly Property [Continue] As Primitive = "Continue"

    End Class

End Namespace
