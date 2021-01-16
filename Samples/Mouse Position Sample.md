# Absolute and Relative Mouse Position Sample:
While mouse moves on the form:
- Form1 will show (on its titlebar) ths mouse pos relative to the form.
- TextBoxe1 will show the mouse position relative to itself.
- TextBoxe2 will show the mouse position relative to itself.
- Label1 will show the mise poition relative to the screen, which is the absolute mouse pos.

# Code:
Use the form designer to add 2 textBoxes and a label, then switch to the Form code tab, and add this code after the generated SB code:
```VB.NET
Form1.HandleEvents()
Control.OnMouseMove = TextBox1_MouseMove

Sub TextBox1_MouseMove
  TextBox1.Text = "Pos relative to TextBox1: (" + TextBox1.MouseX + "," + TextBox1.MouseY + ")"
  TextBox2.Text = "Pos relative to TextBox2: (" + TextBox2.MouseX + "," + TextBox2.MouseY + ")"
  Label1.Text = "Pos relative to Screen: (" + Mouse.MouseX + "," + Mouse.MouseY + ")"
  Form1.Text = "Pos relative to Form1: (" + Form1.MouseX + "," + Form1.MouseY + ")"
EndSub
```