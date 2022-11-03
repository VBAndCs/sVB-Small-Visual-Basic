Imports Microsoft.SmallVisualBasic.Library

Namespace WinForms

    ''' <summary>
    ''' Contains the names of winforms controls
    ''' </summary>
    <SmallVisualBasicType>
    Public NotInheritable Class ControlTypes
        ''' <summary>Control</summary>
        <ReturnValueType(VariableType.ControlType)>
        Public Shared ReadOnly Property Control As Primitive = "Control"

        ''' <summary>Form</summary>
        <ReturnValueType(VariableType.ControlType)>
        Public Shared ReadOnly Property Form As Primitive = "Form"

        ''' <summary>TextBox</summary>
        <ReturnValueType(VariableType.ControlType)>
        Public Shared ReadOnly Property TextBox As Primitive = "TextBox"

        ''' <summary>Label</summary>
        <ReturnValueType(VariableType.ControlType)>
        Public Shared ReadOnly Property Label As Primitive = "Label"

        '''' <summary>ImageBox</summary>
        '<ReturnValueType(VariableType.ControlType)>
        'Public Shared ReadOnly Property ImageBox As Primitive = "ImageBox"

        ''' <summary>ListBox</summary>
        <ReturnValueType(VariableType.ControlType)>
        Public Shared ReadOnly Property ListBox As Primitive = "ListBox"

        ''' <summary>ComboBox</summary>
        <ReturnValueType(VariableType.ControlType)>
        Public Shared ReadOnly Property ComboBox As Primitive = "ComboBox"

        ''' <summary>DatePicker</summary>
        <ReturnValueType(VariableType.ControlType)>
        Public Shared ReadOnly Property DatePicker As Primitive = "DatePicker"

        ''' <summary>CheckBox</summary>
        <ReturnValueType(VariableType.ControlType)>
        Public Shared ReadOnly Property CheckBox As Primitive = "CheckBox"

        ''' <summary>RadioButton</summary>
        <ReturnValueType(VariableType.ControlType)>
        Public Shared ReadOnly Property RadioButton As Primitive = "RadioButton"

        ''' <summary>Button</summary>
        <ReturnValueType(VariableType.ControlType)>
        Public Shared ReadOnly Property Button As Primitive = "Button"

        ''' <summary>MenuItem</summary>
        <ReturnValueType(VariableType.ControlType)>
        Public Shared ReadOnly Property MenuItem As Primitive = "MenuItem"

        ''' <summary>Menu</summary>
        <ReturnValueType(VariableType.ControlType)>
        Public Shared ReadOnly Property MainMenu As Primitive = "Menu"

    End Class

End Namespace
