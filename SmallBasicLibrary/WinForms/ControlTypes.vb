Imports Microsoft.SmallVisualBasic.Library

Namespace WinForms

    ''' <summary>
    ''' Contains the names of winforms controls
    ''' </summary>
    <SmallVisualBasicType>
    Public NotInheritable Class ControlTypes
        ''' <summary>Control</summary>
        <ReturnValueType(VariableType.ControlType)>
        Public Shared ReadOnly Property Control As New Primitive("Control")

        ''' <summary>Form</summary>
        <ReturnValueType(VariableType.ControlType)>
        Public Shared ReadOnly Property Form As New Primitive("Window")

        ''' <summary>TextBox</summary>
        <ReturnValueType(VariableType.ControlType)>
        Public Shared ReadOnly Property TextBox As New Primitive("TextBox")

        ''' <summary>Label</summary>
        <ReturnValueType(VariableType.ControlType)>
        Public Shared ReadOnly Property Label As New Primitive("Label")

        '''' <summary>ImageBox</summary>
        '<ReturnValueType(VariableType.ControlType)>
        'Public Shared ReadOnly Property ImageBox As Primitive = "ImageBox"

        ''' <summary>ListBox</summary>
        <ReturnValueType(VariableType.ControlType)>
        Public Shared ReadOnly Property ListBox As New Primitive("ListBox")

        ''' <summary>ComboBox</summary>
        <ReturnValueType(VariableType.ControlType)>
        Public Shared ReadOnly Property ComboBox As New Primitive("ComboBox")

        ''' <summary>DatePicker</summary>
        <ReturnValueType(VariableType.ControlType)>
        Public Shared ReadOnly Property DatePicker As New Primitive("DatePicker")

        ''' <summary>CheckBox</summary>
        <ReturnValueType(VariableType.ControlType)>
        Public Shared ReadOnly Property CheckBox As New Primitive("CheckBox")

        ''' <summary>RadioButton</summary>
        <ReturnValueType(VariableType.ControlType)>
        Public Shared ReadOnly Property RadioButton As New Primitive("RadioButton")

        ''' <summary>Button</summary>
        <ReturnValueType(VariableType.ControlType)>
        Public Shared ReadOnly Property Button As New Primitive("Button")

        ''' <summary>MenuItem</summary>
        <ReturnValueType(VariableType.ControlType)>
        Public Shared ReadOnly Property MenuItem As New Primitive("MenuItem")

        ''' <summary>Menu</summary>
        <ReturnValueType(VariableType.ControlType)>
        Public Shared ReadOnly Property MainMenu As New Primitive("Menu")

        ''' <summary>Slider</summary>
        <ReturnValueType(VariableType.ControlType)>
        Public Shared ReadOnly Property Slider As New Primitive("Slider")

        ''' <summary>ProgressBar</summary>
        <ReturnValueType(VariableType.ControlType)>
        Public Shared ReadOnly Property ProgressBar As New Primitive("ProgressBar")

        ''' <summary>ScrollBar</summary>
        <ReturnValueType(VariableType.ControlType)>
        Public Shared ReadOnly Property ScrollBar As New Primitive("ScrollBar")


        ''' <summary>WinTimer</summary>
        <ReturnValueType(VariableType.ControlType)>
        Public Shared ReadOnly Property WinTimer As New Primitive("WinTimer")

        ''' <summary>ToggleButton</summary>
        <ReturnValueType(VariableType.ControlType)>
        Public Shared ReadOnly Property ToggleButton As New Primitive("ToggleButton")

    End Class

End Namespace
