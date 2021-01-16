# Change Trancparenct sample:
Eech time a butten is clicked, encreases the trancparency of the TextBox backcolor by 10%.

# Code:
Use the form designer to add a textBox and a button, switch to the Form code tab, and add this code after the generated SB code:
```VB.NET
Button1.HandleEvents()
Control.OnMouseLeftUp = Click
TextBox1.BackColor = Color.Red
Trans = 0

Sub Click
  Trans = Trans + 10
  TextBox1.BackColor = Color.SetTransparency(Color.Red, Trans)
  TextBox1.Text = Color.GetNameAndTransparency(TextBox1.BackColor)
EndSub
```


