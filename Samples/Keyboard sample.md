# Keyboard and Keys modules sample:
Every time ypu press a key from keyboard, the label will show that stae of modifier keys and the name pf the key that is pressed and its num in Keys enum.
If ypu press the Escabe key, the form will be closed.

# Code:
Use the form designer to add a textbox and a label to the form and switch to the Form Code tab, and insert the following code after the generated code:
```VB.NET
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
```  