Imports Microsoft.SmallVisualBasic.Library

Namespace WinForms

    ''' <summary>
    ''' Provides methods to infer the given value as a specific VariableType.
    ''' </summary>
    <SmallVisualBasicType> Public NotInheritable Class Convert

        ''' <summary>
        ''' Makes sVB infer the given value as a string.
        ''' </summary>
        ''' <param name="value">The value you need to infer its type.</param>
        ''' <returns>The string representaion of the value.</returns>
        <ReturnValueType(VariableType.String)>
        Public Shared Shadows Function ToStr(value As Primitive) As Primitive
            Return value.AsString()
        End Function

        ''' <summary>
        ''' Makes sVB infer the given value as a double.
        ''' </summary>
        ''' <param name="value">The value you need to infer its type.</param>
        ''' <returns>The numeric representaion of the value.</returns>
        <ReturnValueType(VariableType.Double)>
        Public Shared Function ToDouble(value As Primitive) As Primitive
            Return value.AsDecimal()
        End Function

        ''' <summary>
        ''' Makes sVB infer the given value as a boolean.
        ''' </summary>
        ''' <param name="value">The value you need to infer its type.</param>
        ''' <returns>True or False based on the given value.</returns>
        <ReturnValueType(VariableType.Boolean)>
        Public Shared Function ToBoolean(value As Primitive) As Primitive
            Return CBool(value)
        End Function

        ''' <summary>
        ''' Makes sVB infer the given value as a date.
        ''' </summary>
        ''' <param name="value">The value you need to infer its type.</param>
        ''' <returns>The date representaion of the given value if it can be converted to date, otherwise an empty striing.</returns>
        <ReturnValueType(VariableType.Date)>
        Public Shared Function ToDate(value As Primitive) As Primitive
            Dim d = value.AsDate()
            Return If(d, New Primitive(""))
        End Function

        ''' <summary>
        ''' Makes sVB infer the given value as an array.
        ''' </summary>
        ''' <param name="value">The value you need to infer its type.</param>
        ''' <returns>The same exact value, but sVB will infer it as an array.</returns>
        <ReturnValueType(VariableType.Array)>
        Public Shared Function ToArray(value As Primitive) As Primitive
            Return value
        End Function

        ''' <summary>
        ''' Makes sVB infer the given value as a sound.
        ''' </summary>
        ''' <param name="value">The value you need to infer its type.</param>
        ''' <returns>The same exact value, but sVB will infer it as a sound.</returns>
        <ReturnValueType(VariableType.Sound)>
        Public Shared Function ToSound(value As Primitive) As Primitive
            Return value
        End Function

        ''' <summary>
        ''' Makes sVB infer the given value as a shape.
        ''' </summary>
        ''' <param name="value">The value you need to infer its type.</param>
        ''' <returns>The same exact value, but sVB will infer it as a shape.</returns>
        <ReturnValueType(VariableType.Shape)>
        Public Shared Function ToShape(value As Primitive) As Primitive
            Return value
        End Function

        ''' <summary>
        ''' Makes sVB infer the given value as a color.
        ''' </summary>
        ''' <param name="value">The value you need to infer its type.</param>
        ''' <returns>The same exact value, but sVB will infer it as a color.</returns>
        <ReturnValueType(VariableType.Color)>
        Public Shared Function ToColor(value As Primitive) As Primitive
            Return value
        End Function

        ''' <summary>
        ''' Makes sVB infer the given value as a key.
        ''' </summary>
        ''' <param name="value">The value you need to infer its type.</param>
        ''' <returns>The same exact value, but sVB will infer it as a key.</returns>
        <ReturnValueType(VariableType.Key)>
        Public Shared Function ToKey(value As Primitive) As Primitive
            Return value
        End Function

        ''' <summary>
        ''' Makes sVB infer the given value as a dialog result.
        ''' </summary>
        ''' <param name="value">The value you need to infer its type.</param>
        ''' <returns>The same exact value, but sVB will infer it as a dialog result.</returns>
        <ReturnValueType(VariableType.DialogResult)>
        Public Shared Function ToDialogResult(value As Primitive) As Primitive
            Return value
        End Function

        ''' <summary>
        ''' Makes sVB infer the given value as a control type.
        ''' </summary>
        ''' <param name="value">The value you need to infer its type.</param>
        ''' <returns>The same exact value, but sVB will infer it as a control type.</returns>
        <ReturnValueType(VariableType.ControlType)>
        Public Shared Function ToControlType(value As Primitive) As Primitive
            Return value
        End Function

        ''' <summary>
        ''' Makes sVB infer the given value as a control.
        ''' </summary>
        ''' <param name="value">The value you need to infer its type.</param>
        ''' <returns>The same exact value, but sVB will infer it as a control.</returns>
        <ReturnValueType(VariableType.Control)>
        Public Shared Function ToControl(value As Primitive) As Primitive
            Return value
        End Function

        ''' <summary>
        ''' Makes sVB infer the given value as a form.
        ''' </summary>
        ''' <param name="value">The value you need to infer its type.</param>
        ''' <returns>The same exact value, but sVB will infer it as a form.</returns>
        <ReturnValueType(VariableType.Form)>
        Public Shared Function ToForm(value As Primitive) As Primitive
            Return value
        End Function

        ''' <summary>
        ''' Makes sVB infer the given value as a text box.
        ''' </summary>
        ''' <param name="value">The value you need to infer its type.</param>
        ''' <returns>The same exact value, but sVB will infer it as a text box.</returns>
        <ReturnValueType(VariableType.TextBox)>
        Public Shared Function ToTextBox(value As Primitive) As Primitive
            Return value
        End Function

        ''' <summary>
        ''' Makes sVB infer the given value as a label.
        ''' </summary>
        ''' <param name="value">The value you need to infer its type.</param>
        ''' <returns>The same exact value, but sVB will infer it as a label.</returns>
        <ReturnValueType(VariableType.Label)>
        Public Shared Function ToLabel(value As Primitive) As Primitive
            Return value
        End Function

        ''' <summary>
        ''' Makes sVB infer the given value as an image box.
        ''' </summary>
        ''' <param name="value">The value you need to infer its type.</param>
        ''' <returns>The same exact value, but sVB will infer it as an image box.</returns>
        <ReturnValueType(VariableType.ImageBox)>
        Public Shared Function ToImageBox(value As Primitive) As Primitive
            Return value
        End Function

        ''' <summary>
        ''' Makes sVB infer the given value as a list box.
        ''' </summary>
        ''' <param name="value">The value you need to infer its type.</param>
        ''' <returns>The same exact value, but sVB will infer it as a list box.</returns>
        <ReturnValueType(VariableType.ListBox)>
        Public Shared Function ToListBox(value As Primitive) As Primitive
            Return value
        End Function

        ''' <summary>
        ''' Makes sVB infer the given value as a combo box.
        ''' </summary>
        ''' <param name="value">The value you need to infer its type.</param>
        ''' <returns>The same exact value, but sVB will infer it as a combo box.</returns>
        <ReturnValueType(VariableType.ComboBox)>
        Public Shared Function ToComboBox(value As Primitive) As Primitive
            Return value
        End Function

        ''' <summary>
        ''' Makes sVB infer the given value as a date picker.
        ''' </summary>
        ''' <param name="value">The value you need to infer its type.</param>
        ''' <returns>The same exact value, but sVB will infer it as a date picker.</returns>
        <ReturnValueType(VariableType.DatePicker)>
        Public Shared Function ToDatePicker(value As Primitive) As Primitive
            Return value
        End Function

        ''' <summary>
        ''' Makes sVB infer the given value as a check box.
        ''' </summary>
        ''' <param name="value">The value you need to infer its type.</param>
        ''' <returns>The same exact value, but sVB will infer it as a check box.</returns>
        <ReturnValueType(VariableType.CheckBox)>
        Public Shared Function ToCheckBox(value As Primitive) As Primitive
            Return value
        End Function

        ''' <summary>
        ''' Makes sVB infer the given value as a radio button.
        ''' </summary>
        ''' <param name="value">The value you need to infer its type.</param>
        ''' <returns>The same exact value, but sVB will infer it as a radio button.</returns>
        <ReturnValueType(VariableType.RadioButton)>
        Public Shared Function ToRadioButton(value As Primitive) As Primitive
            Return value
        End Function

        ''' <summary>
        ''' Makes sVB infer the given value as a toggle button.
        ''' </summary>
        ''' <param name="value">The value you need to infer its type.</param>
        ''' <returns>The same exact value, but sVB will infer it as a toggle button.</returns>
        <ReturnValueType(VariableType.ToggleButton)>
        Public Shared Function ToToggleButton(value As Primitive) As Primitive
            Return value
        End Function

        ''' <summary>
        ''' Makes sVB infer the given value as a button.
        ''' </summary>
        ''' <param name="value">The value you need to infer its type.</param>
        ''' <returns>The same exact value, but sVB will infer it as a button.</returns>
        <ReturnValueType(VariableType.Button)>
        Public Shared Function ToButton(value As Primitive) As Primitive
            Return value
        End Function

        ''' <summary>
        ''' Makes sVB infer the given value as a menu item.
        ''' </summary>
        ''' <param name="value">The value you need to infer its type.</param>
        ''' <returns>The same exact value, but sVB will infer it as a menu item.</returns>
        <ReturnValueType(VariableType.MenuItem)>
        Public Shared Function ToMenuItem(value As Primitive) As Primitive
            Return value
        End Function

        ''' <summary>
        ''' Makes sVB infer the given value as a main menu.
        ''' </summary>
        ''' <param name="value">The value you need to infer its type.</param>
        ''' <returns>The same exact value, but sVB will infer it as a main menu.</returns>
        <ReturnValueType(VariableType.MainMenu)>
        Public Shared Function ToMainMenu(value As Primitive) As Primitive
            Return value
        End Function

        ''' <summary>
        ''' Makes sVB infer the given value as a progress bar.
        ''' </summary>
        ''' <param name="value">The value you need to infer its type.</param>
        ''' <returns>The same exact value, but sVB will infer it as a progress bar.</returns>
        <ReturnValueType(VariableType.ProgressBar)>
        Public Shared Function ToProgressBar(value As Primitive) As Primitive
            Return value
        End Function

        ''' <summary>
        ''' Makes sVB infer the given value as a slider.
        ''' </summary>
        ''' <param name="value">The value you need to infer its type.</param>
        ''' <returns>The same exact value, but sVB will infer it as a slider.</returns>
        <ReturnValueType(VariableType.Slider)>
        Public Shared Function ToSlider(value As Primitive) As Primitive
            Return value
        End Function

        ''' <summary>
        ''' Makes sVB infer the given value as a scroll bar.
        ''' </summary>
        ''' <param name="value">The value you need to infer its type.</param>
        ''' <returns>The same exact value, but sVB will infer it as a scroll bar.</returns>
        <ReturnValueType(VariableType.ScrollBar)>
        Public Shared Function ToScrollBar(value As Primitive) As Primitive
            Return value
        End Function

        ''' <summary>
        ''' Makes sVB infer the given value as a win timer.
        ''' </summary>
        ''' <param name="value">The value you need to infer its type.</param>
        ''' <returns>The same exact value, but sVB will infer it as a win timer.</returns>
        <ReturnValueType(VariableType.WinTimer)>
        Public Shared Function ToWinTimer(value As Primitive) As Primitive
            Return value
        End Function

    End Class
End Namespace
