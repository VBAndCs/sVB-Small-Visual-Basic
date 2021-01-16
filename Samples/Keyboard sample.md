' A sample on using the Keyboard and Keys modules
' Use the form designer to add a textbox and a label to the form
' and switch to the Form Code tab,
' and insert the following code after the generated code:

TextBox1.HandleEvents()
Control.OnKeyUp = TextBox1_KeyUp

Sub TextBox1_KeyUp
  info = ""
  if KeyBoard.AltPressed Then
    info = info + "Alt+"
  EndIf
  
  if KeyBoard.CtrlPressed Then
    info = info + "Ctrl+"
  EndIf
  
  if KeyBoard.ShiftPressed Then
    info = info + "Shift+"
  EndIf
  
  info = info + Keyboard.LastKeyName + ": " + Keyboard.LastKey
  
  If Keyboard.LastKey = Keys.Escape Then
     Form.Close(Form1)   
  EndIf
 
  Label1.Text = info
EndSub
  