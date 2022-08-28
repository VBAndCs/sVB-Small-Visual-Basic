
# Contents
- Small Visual Basic (sVB):
- Using the source code:
- Download the language:
- Try the samples:
- Why do we need sVB:
- Form designer Features:
- SB Code Enhancements:
- ToDo:

# What's new in sVB v1.6
1. Some bug fixes in compiler and other tools.
2. Add `OnShown` event to the Form. It is fired after the content of the form is rendered. If you use it, you should add all initialization into it, as you can't know for sure if it will occur before or after the code in the global section is executed! But you can be sure that all the controls are rendered and ready for use.
`OnShown` is the default event for the Form, and you can add a handler to it by just double-clicking the form surface in the form designer.
3. I got rid of the side help panel to save space, and showed the help info in a tip window that pops up after 2 seconds from moving the caret to any word in the code editor, and stays open for 10 seconds unless you move to another position, move the scroll, or press Esc key. I prevent showing the help for the same word unless toy move to another one, but you can force to show the help by pressing `F1`.
You can magnify the font of the popup help by pressing Ctrl and moving the mouse wheel. This is one of many functionalites built-in the WPF FlowDocument control that is used to show the help.
You can say I brought the VS intellisense to sVB. The pop up help offers valuable info about the current code token, including:
 * The scope (local or global var).
 * The definition signature (Type, Property, Dynamic Property, Event, Method Parameters).
 * If the token is a `variable`, a `sub` or a `function`, it will be shown as a link, so, you can click it to go to its definition line. If the token it the name of the form or a control on it, clicking the link will select the form or the control on the form designer.
* The documentation includes a summery, and info about parameters and return value. You can add a summery for user defined types by adding one or more comment lines above the var, sub, or function definitions. You can also add one more comment line at the end of the definition line. For subs and functions, you can add the additional summery line after the opining parans if you split the params over multi lines which also allows you to add a comment for each parameter to be used as a documentation. For Functions, the comment placed after the closing parans will be used as the documentation info for the return value. For subs, it will be considered an additional line of the summery. Ex:
```VB
XPos = 1   ' the horizontal position 

' adds x to the pos
Sub Move(
    x ' The increment value to add to the pos
)
   XPos = XPos + x
EndSub

Function InRange( ' Checks if the pos is withing the givin range
    start, ' the start position
    end  ' the end position
) ' True if the pos is in range, False otherwise.
   Return XPos >= start And XPos <= end
EndFunction
```

![image](https://user-images.githubusercontent.com/48354902/187048217-0f626439-8e90-4e90-b838-41cd3312ae4b.png)

While typing the arguments of the method call, the popup help will highlight the current param with a red color, and show only the info about this param, so you can focus only on the task in hand.

![image](https://user-images.githubusercontent.com/48354902/187048339-8861202b-9b86-41ce-9af9-b09c4dc3d7d1.png)

4. The editor formats the current sub after leaving a line that has changes. Formatting doesn't only include indentation, but also pretty listing of space between tokens, and fixing the casing of keywords, labels, type and method names. It also enforces using lower-case initial letters for local variables, and upper-case initial letters for global variables, labels, subs and functions.

5. The editor highlights every occurrence of the current identifier (variable, sub, function, label, type, method) name. Similar to highlighted block keywords, you can navigate between highlighted identifiers by pressing `F4` or `Ctrl+Shift+Up` or `Ctrl+Shift+down`

6. Many enhancements to the completion list to make it smarter, such as :
 * Filtering completion names by partial words (for ex: typing name can select MyName) 
 * Filtered out names don't appear in the list anymore.
 * The list remembers last selected object for each first letter.
 * The list remembers last selected method for each object.

The aim of this release is to make coding with sVB easier, educational and fun!

# Small Visual Basic (sVB):
sVB is an evolved version of Microsoft Small Basic with a small WinForms library and a graphics form designer. 

![Untitled](https://user-images.githubusercontent.com/48354902/126494834-f6c90190-4241-40c7-84b3-c3b3c432a6d1.png)

sVB has many enhancements over SB to make writing apps fast and easy  with little code. It brings back the joy and excitement of using vb6 to write RAD applications, by adding the illusion of the Object type while accessing controls members:
```vb
Label1.Text = TextBox1.Text + TextBox2.Text
```

All Label1, TextBox1 and TextBox2 are just string variables, but in sVB you can use the names of the controls as if they are Objects containing the controls themselves, and hence access their properties, methods and events directly, with support of the auto-completion list.

![untitled2](https://user-images.githubusercontent.com/48354902/126494901-60dfa36b-cdaf-4fd0-8107-1769f4e5c4c4.jpg)

You can also add event handlers from the upper dropdown lists: 
Choose the control name from the left list (say `Button1`), and click the event name from the right list (say `OnClick`), and this sub will be added for you in the code editor:
```vb
Sub Button1_OnClick
   
EndSub
```

Or you can simply double-click the button on the form designer, and the Button1_OnClick will be created for you!

To make this work, each form created by the designer has 3 files:
1. a `.xaml` file containing the form design.
2. a `.sb.gen` file, containing normal SB code that defines vars to hold form and controls names, and adds the event handlers instructions to connect each event with the sub that handles it.
3. a `.sb` file that you write your code in without warring about the other 2 files contents, as they are fully generated by the designer. This makes you write short, simple and clean code, focusing only on the task in hand. See the [samples folder](https://github.com/VBAndCs/sVB-Small-Visual-Basic/tree/master/Samples).

# Using the source code:
sVB is fully written with VB.NET, and targets .NET framework 4.8. You can run the source code in VS.NET 2019 and later.
befor running the code, please copy the two folders `Lib` and `Toolbar` from the `SmallBasicIDE\SB.Lib` folder to the `SmallBasicIDE\bin\Debug` folder, ass obviously git execluds this folder, and I prefer it this way.

# Download the language:

Go to the [Releases page](https://github.com/VBAndCs/sVB-Small-Visual-Basic/releases), navigate to the latest version of vSB, expand the Assets list at the bottom of the page, and download the ZIP file.
Follow these instructions:
1. sVB needs .NET framework 4.8. If you don't have it on your PC, download and install it. https://go.microsoft.com/fwlink/?LinkId=2085155

2. Unzip the `sVB.zip` file. You will have a folder with the same name where you unzipped the file. Open the folder and double-click `sVB.exe`.
And that it. You are ready to go :)

# Try the samples:
1. Right-click the form designer and click `Open` from the context menu. 
2. In the `open file dialog`, Navigate to the `sVB\Samples` folder and open any sample folder. 
3. Select the `.xaml` file from the folder and click the `Open` button. This will open the form in the designer.
4. Click the `Form code` tab at the top of the window to switch to the sb code editor.
5. Click the `Run` button (or hit F5 from keyboard) to run the program.

You can also open the sample folder in Windows Explorer, right-click the `.sb` and chose `open with` from the context menu, and choose sVB.exe as the default program to open `.sb` files. After that you can just double-click any `.sb` file to open it in sVB.

# Why do we need sVB:
BASIC is famous of being easy to learn, because its syntax is simple  and close to natural English. 
In 2008, MS released Small Basic for kids of 7 years old and above. It was really small, containing only 14 keywords to perform the basic programming instructions like `Sub`, `If`, `For`, `While` and `Goto` statements.

Small basic is a dynamic language, as it doesn't declare variables with types. You just assign a value to a valid identifier and SB will declare it as a Primitive variable, which can hold a string, a number, or an array.
This makes the language very easy to learn and use for kids.

But, a programming language is not just instructions. It has to have a class library to communicate with Windows. In fact SB comes with a number of powerful libraries, and allow you to supply your own libraries as well. This is where I saw a big issue while trying to teach SB to my nephews:
The language is too easy,  but using the libraries isn't that easy, especially when dealing with graphics and UI. The PDF book that comes with SB makes it hard, focusing on drawing shapes by using geometric functions (it even contains a fractals sample!)  
This is not the best way to introduce programming to kids. A black command window (the TextWindow) is easy but boring, while using vector graphics or drawing using the Turtle on the Graphics Window is amazing but can be quite hard.

The good news is that the Controls class allows you to draw a TextBox, and a Button on the GraphicsWindow, deal with their properties and handle their events. But, unfortunately, the kid has to design the form blindly while adding controls by code, and even worst,  the code used to communicate with these controls is verbose, because SB doesn't have the Object type as I explained above, so, you can only store the name of the control in a variable:
```vb
btn = Controls.AddButton("Enable", 100, 100)
Controls.ButtonClicked = OnClick
```

then send this variable to each method you use to alter the control:
```VB
Sub OnClick
   If Controls.GetButtonCaption(btn) = "Enable" Then
      Controls.SetButtonCaption(btn, "Disable")
   Else
      Controls.SetButtonCaption(btn, "Enable")
   EndIf
EndSub
```

This is not the kind of code you want to show to a kid!
In fact it will be easier to teach Visual Basic to the kid, so he can drag a button form the toolbox, drop it on the window, set it's name and caption from the properties window, double-click it to go to it's click event handler in the code editor, and just write:
```vb
   If btn.Text = "Enable"
      btn.Text = "Disable"
   Else
      btn.Text = "Enable"
   EndIf
```

And that's it. Fast, clean, easy and  short code, that made us love programming!

It is unbelievable that SB complicated such an easy task, in the name of being simple and easy to learn for kids!

I looked at some SB alternative IDEs but they are either:
- more complex (too advanced to do nothing important with a language meant to be a leering toy),
- or simple enough to draw the controls and generate some code for them, but still can't overcome the SB syntax limitations when dealing with objects.

This is when I decides to do something, and here we are.

# Form designer Features:
I will write this later, but the form designer is too easy to use, so, enjoy trying it.

# SB Code Enhancements:
I made many improvements to the SB compiler:
1. Support array initializers:

You can use the `{}` to set multiple elements to the array at once:
```vb
x = {1, 2, 3}
```

Nested initializers are also supported when you deal with multi-dimensional arrays:
```vb
y = {"a", "b", {1, 2, 3}}
```

You can also use vars inside the initializer, so, the above code can be rewritten as:
```vb 
y = {"a", "b", x}
```

And you can send an initializer as a param to a function:
```vb
TextWindow.WriteLine({"Adam", 11, "Succeeded"})
```

2. `For Next` and `While Wend`:

SB uses `EndFor` and `EndWhile` to close `For` and `While` blocks respectively. This is still supported in sVB but I allowed also to use `Next` to close `For` and `Wend` to close `While`, as they are used in VB6. I encourage you to use Next and Wend, as they give the meaning of repeating and circulating over the loop. `End` gives the meaning of finishing and exiting, so, it is suitable in `EndIf` and `EndSub`, but confusing in `EndFor` and `EndWhile`, as they can imply that `the loop finishes here`, not just `the block ends here`!

3. You can use `ExitLoop` to exit For and While loops, and `ContinueLoop` to skip the current iteration and jump back to the beginning of the loop body to continue the next iteration in for loop. Be aware that unlike For, while doesn't have an auto-incremented counter, so be sure you write the suitable code to update the variable that while condition depends on before using ContinueLoop inside the while block, otherwise you may end up stuck with an infinite loop.
In the case of nested 2 loops of any kind, you can exit the outer loop by using `ExitLoop -`. You can add more `-` signs to exit up levels loops in case you have 3 or 4 nested loops, or just use `ExitLoop *` to exit all nested loops at once. The same rules applies to `ContinueLoop` if you want to use it to cointinue outer loops.

4. You can use `Me` to refer to the current Form.

5. 'True' and 'False' are keywords of sVB.

6. Subroutines can have parameters now:
```vb
Sub Print(Name, Value)
   TextWindow.WriteLine("Name=" + Name + ", Value=" + Value)
EndSub
```

And call this sub like this:
```vb
  Print("Distance", 120)
```

Note that you can use `Return` inside the sub body to exit the sub immediately.


7. sVB can define functions now. You can supply params to get the fyunction input and use `Return` to return the function output.
```vb
Function Sum(x, y)
    Return x + Y
EndFunction
```

And you can use it like this:
```vb
x = Sum(1, 2)
```

8. SB doesn't have variable scope, as all variables are considered global, and you can define them in any place in the file and use them from any other place in the file (up or down). sVB has cleaned this mess, which is a break change that can prevent some SB code from running probably in sVB, but it is a necessary step to make the kid organize his code and write clean code. This is also necessary to make sub and function params work correctly, and allow you to use recursive subs and functions. The mew scope rules are:
- Sub and function params are local, and hides any global vars with the same names.
- The For loop counter(iterator) in local and hodes any global var with the same name.
- Any var defined inside the sub or the function is local unless there is a global var with the same name is defined above of the sub function. If the global var is defined below, then the local var will hide it.

So, as a good practice:
- Define all global vars at the very top pf the file.
- Give global vars a prefix (such as `g_`) to avoid any conflictions with local vars.
- Don't use global vars unless there is no other solution, instead pass values to subs and functions through params, and receive values from functions through their return values.

9. The editor auto completes If, For, While, and Sub blocks just after writing a space after them.

10. The editor has a perfect auto-indentation.

11. Using naming convention to give sVB some info about the var type to make using them easier. 
- The first naming convention: Any var ends with or starts with a control name (like Form1 or myLabel) will be treated as a Label, so you can use the short syntax Control.Property and the auto completion list will appear to help you complete method and properties names.

- The second naming convention: Any var ends with or starts with the word Data is treated as a dynamic object, and you can add properties to it, and get auto completion support when you use them.
```
CarData.Color = Color.Red
CarData.Speed = 100
x = CarData.Speed
```

In fact, sVB converts the above syntax to:
```
CarData["Color"] = Color.Red
CarData["Speed"] = 100
x = CarData["Speed"]
```

Note that this naming convention rule ignores var domain rules, to allow you reuse the properties across subs and functions. This is totally safe as it only affects the property list that will appear in the auto completion list, but it has no effect on the variable domain rules when you run the program. You may see properties names from a data object from another function, but you still can't read these properties values in code. It is just a way to make coding faster and easier.
- The third naming convention: Any Data var that contains the name of another data var (after trimming the `Data` part of them bath) will show the union of their properties in the auto completion list. This allows you to `inherit` other data properties. As an example, of you use the names Car2Data, or myCarData in the above example, they will show the Color and Speed properties (coming from CarData) in the completion list"
```
Car2Data.Speed = 200
Car2Data.Acceleration = 10
```

And if you use `MyCar2Data` you will inherit all properties from `MyCarData`, `Car2Data` and `CarData`!

12. Use the vb lookup operator to crate dynamic properties:

```VB
Student!ID = 1
Student!Name = "Adam"
```

This is a shorter alternative to using the Data naming convention (see more details at the end of the file):
```VB
StudentData.ID = 1
StudentData.Name = "Adam"
```

13. Many enhancements in the WinForms controls and bug fixes.

14. You can split the code line over multi-lines by using the _ symbol at the end of each  
line segment. ex:
```VB..NET
Dim x = "First Line" + _
            "Second Line"
```

You can use _ at any position expet before of after the dot `.`.

15. You can also split the line at some positions without using the _ . These positions are:
- after the following symbols: `,`, `=`, `+`, `-`, `*`. `/`, `(`, `[`, `{`, `or`, `and`.
- before the following symbols: `+`, `-`, `*`. `/`, `)`, `]`, `}`, `or`, `and`.

Ex:
```VB
If x = y or 
    Text.GetSubText(
       "some text",
       6,
       3
   ) = "abc" Then

   x = 0
End If
```

16. You can add comments at the end of any line segment. Ex:
```VB
Function Add(  ' Adds two numbers
   a, _   ' first number
   b      ' second number
)
   Return a + b 
EndFunction
```

17. Adding an `(Add new function)` command to the upper right dropdown list in the code  
editor

18. Enhancing the auto completion list.

19. Enhancing the auto formatting of code, to adjust indentation of sub lines, and adjust  
spaces between tokens.

20. The code editor now highlights any matching pairs such as `()`, `[]` and `{}`. It also  
highlights the keywords of the Sub, Function, If, For, and While bolcks whenever you insert  
the caret on on of them. You can move to the next highlighted token by pressing `F4` or  
`Ctrl+Shift+Down`, and you can move to the previous highlighted token by pressing `Shift 
+F4` or `Ctrl+Shift+Up`.
You can also press such shortcuts keys on any line even there is no highlighted keywords.  
This will highlight the nearest block that contains the statement, and move to the neareest  
keyword up or down according to the shourtcut keys you are using.

# ToDo:
- Add more controls to the winForms library.
- Support multi-from programs. Right now you can define other forms by code, but this will bring you back to the verbose code similar to SB's controls code. The designer can already open multiple forms, but I need to add the concept of the Project, and change SB compiler to add multiple modules in the Exe, so it can compile multiple form files into one exe.
