<p align="center">
  <img src="https://github.com/VBAndCs/sVB-Small-Visual-Basic/assets/48354902/9e1ebcbf-212f-4578-8b67-b57551b80270"/>
</p>

# Contents
* [What is Small Visual Basic?](https://github.com/VBAndCs/sVB-Small-Visual-Basic#what-is-small-visual-basic)
* [Download the language](https://github.com/VBAndCs/sVB-Small-Visual-Basic#download-the-language)
* [sVB full reference PDF book](https://github.com/VBAndCs/sVB-Small-Visual-Basic#svb-full-reference-pdf-book)
* [Try the samples](https://github.com/VBAndCs/sVB-Small-Visual-Basic#try-the-samples)
* [Using the sVB source code](https://github.com/VBAndCs/sVB-Small-Visual-Basic#using-the-svb-source-code)
* [Why do we need sVB?](https://github.com/VBAndCs/sVB-Small-Visual-Basic#why-do-we-need-svb)
* [It is also a Small Visual Basic .NET!](https://github.com/VBAndCs/sVB-Small-Visual-Basic#it-is-also-a-small-visual-basic-net)
* [sVB for kids](https://github.com/VBAndCs/sVB-Small-Visual-Basic#svb-for-kids)
* [Conclusions](https://github.com/VBAndCs/sVB-Small-Visual-Basic#conclusions)


# What is Small Visual Basic?
Small Visual Basic (sVB) is an educational programming language, created by Eng. Mohammad Hamdy as an evolved version of Microsoft Small Basic (SB). It is meant to be easier and more powerful at the same time, to introduce programming basics to kids and beginners of any age, provided that they can use the English keyboard on the Windows operating system.

# Download the language:
Starting from v2.8.1, sVB has an installer, that creates a startup menu and desktop shortcuts for sVB and its book.
You can download the installer from the [visualstudio marketplace](https://marketplace.visualstudio.com/items?itemName=ModernVBNET.sVBInstaller)
To install sVB, Unzip the downlaoded file and run the sVB_Installer.msi. You can accept default options and click next in each setup page, and wait for the setup to finish. If your PC doesn't contain the .NET framework 4.5, the estup will ask you to download it, so you should accept and install it.
To get older versions, go to the [Releases page](https://github.com/VBAndCs/sVB-Small-Visual-Basic/releases), navigate to the version you want, expand the Assets list at the bottom of the page, and download the ZIP file, and follow these instructions:
1.	sVB needs [.NET framework 4.5](https://go.microsoft.com/fwlink/?LinkId=2085155). If you don't have it on your PC, download and install it.
2.	Unzip the "sVB.zip" file. You will have a folder with the same name where you unzipped the file. Open the folder and double-click "sVB.exe".

# sVB full reference PDF book
When you install sVB, you will find a `sVB docs.pdf`file in its directory, and you will find a shortcut for it in the Small Visual Basic folder in the start menu, and another one on the desktop.
You can [read this file also in the root folder of the sVB repo](https://github.com/VBAndCs/sVB-Small-Visual-Basic/blob/master/sVB%20Docs.pdf).
This is a 750 pages books, that contains the full details about the sVB IDE, syntax, and class library with full examples and helpful notes about the projects you can find in the sVB samples folder.

# Try the samples:
The sVB folder contains a samples folder that contains various projects. To try any sample, follow these instructions:
1.	Right-click the form designer and click "Open" from the context menu, or click the "Open Existing Form" icon on the toolbar, or press the Ctrl+O shortcut from the keyboard.
2.	In the "open file" dialog, navigate to the "sVB\Samples" folder and open the sample folder you want. 
3.	Select any ".xaml" file from the folder and click the "Open" button. This will close the dialog and open the form in the designer.
4.	Click the "Run" button from the toolbar (or hit F5 from keyboard) to run the program. Note that closing the main form will close the program.

You can also open any sample folder in Windows Explorer, right-click any ".sb" file, choose "open with" from the context menu, and choose sVB.exe as the default program to open ".sb" files.
After that you can just double-click any ".sb" file to open it in sVB.

# Using the sVB source code:
sVB is an open source project, that is published as a [GitHub repo](https://github.com/VBAndCs/sVB-Small-Visual-Basic).
The source code of MS Small Basic is written in C#, but I converted a copy of it to VB.NET and uses it to create sVB. 
I am a professional C# programmer, but it seemed inappropriate to me to write a BASIC-family compiler with a C-family language (Not to mention that VB.NET is my favorite language).
Furthermore, I wanted to use sVB as a training yard for understanding and working with compilers, both for me and for the VB.NET community. This small compiler with small tools, provides a very easy way to understand how compilers are built and how they work, which can make it easier in a next step to work with the VB.NET compiler (Which is a part of [Roslyn](https://github.com/dotnet/roslyn)).
This is a necessary step for the VB community, after MS neglected VB.NET since 2017, and stopped evolving it since 2020. And this is why Anthony D. Green forked Roslyn and start evolving VB.NET under the name [ModVB](https://anthonydgreen.net/2022/08/20/introducing-modvb/) since 2022.
Anthony had participated in designing and creating Roslyn and VB.NET in the first place, but he is still one man, and the community must help him maintain and evolve ModVB, if they really want it to live long and prosper. 
So, here is a small programming language written in VB.NET for you. You can fork it, play with it, and evolve it as you want, and when you feel comfortable with compilers, try to take a look at Roslyn and ModVB.
Note that all sVB projects are WPF projects, that target the .NET framework 4.5. You can run the source code in VS.NET 2019 and later. But before running the code, please copy the "Lib" and "Toolbar" folders from the "SmallBasicIDE\SB.Lib" folder to the "SmallBasicIDE\bin\Debug" and "SmallBasicIDE\bin\Release" folders respectively, as obviously "Git" excludes these folders, and I prefer it this way.

# Why do we need sVB:
> "BASIC used to be on every computer a child touched, but today there's no easy way for kids to get hooked on programming."
> David Brin, "Why Johnny can't code", 2006.

One day, Vijaye Raji has read the above article, so he decided to do something to solve that problem. In his spare time in Microsoft, he created the Small Basic as an educational programming language for 7 years old kids and above, and Microsoft had released it in 2008.
I haven't read this article, nor even knew that Small Basic exists, when I tried to teach VB.NET to my 13 years old nephew, which was not easy for him!
By the way, his name is Mohammad Hamdy Ghanem too, as my brother is named after my father, and his son is named after me ?.
So, I found myself asking a similar question: Why can't Mohammad code?
There are tons of fundamental facts about VB.NET that he needs to learn before producing any thing useful.
He kept asking me about who would use such simple samples in the real world, and he was absolutely right!
I had to build a simple car racing game to get his attention, but things wasn't that easy for him! User controls, objects, and inheritance?! He's absolutely got overwhelmed!
I recall that the classic Visual Basic was much easier to start with, even for someone without any programming background, but VB.NET got complicated over years while it was getting more powerful.
When I discussed this with the VB community, one developer suggested to teach him Small Basic instead, and that was the first time I hear about. 
I tried to do that in the next year with my younger nephews, Ali and Omar, and they already got interested with Small Basic.
Small Basic is really small, containing only 14 keywords to perform the basic programming instructions like "Sub", "If", "For", "While" and "Goto" statements.
Small basic is a dynamic language, as it doesn't declare variables of certain types. You just assign a value to a valid identifier and SB will declare it as a Primitive variable, which can hold a string, a number, or an array.
This makes the language very easy to learn and use for kids.
So, Ali and Omar were happy with SB, until they reached the graphics window drawings that depend on the sine and cosine functions, which were hard to them, because they were still in the 4th grade without any trigonometry background!
The PDF book that comes with SB makes it hard y focusing on drawing shapes by using trigonometric functions (it even contains a fractals sample!)
This is not the best way to introduce programming to kids. A black command window (the Text Window) is easy but boring, while using vector graphics or drawing using the turtle on the Graphics Window is amazing but can be quite hard.
The good news is that the Controls class allows you to draw a TextBox, and a Button on the Graphics Window, deal with their properties and handle their events. But, unfortunately, the kid has to design the form blindly while adding controls by code. Furthermore, the code used to communicate with these controls is verbose, because SB doesn't have objects, so, you can only store the name of the control in a variable, then send it to static/shared methods to change its properties or perform its tasks. For example:
```VB
Btn = Controls.AddButton("Enable", 100, 100)
Controls.ButtonClicked = OnClick

Sub OnClick
   If Controls.GetButtonCaption(Btn) = "Enable" Then
      Controls.SetButtonCaption(Btn, "Disable")
   Else
      Controls.SetButtonCaption(Btn, "Enable")
   EndIf
EndSub
```

This is not the kind of code you want to show to a kid!
In fact it will be easier to teach him Visual Basic, so he can drag a button form the toolbox, drop it on the window, set it's name and caption from the properties window, double-click it to go to it's click event handler in the code editor, and just write:
```VB
If btn.Text = "Enable"
   btn.Text = "Disable"
Else
   btn.Text = "Enable"
EndIf
```

And that's it. A fast, clean, easy and short code, that made us love programming!
It is unbelievable that SB complicated such an easy task, in the name of being simple and easy to learn for kids!
I tried some SB alternative IDEs, but they are either:
-	more complex (too advanced to do nothing important with a language meant to be a learning toy),
-	or simple enough to draw the controls and generate some code for them, but still can't overcome the SB syntax limitations when dealing with objects.
This is why I asked my self: Why does SB have only two windows?
Can't SB have a form designer and a simple syntax to program the controls?
I actually asked for this on the SB repo on GitHub, but got no attention from the team, so, I decided to prove the concept myself.  
I added a form designer to SB, which I built using a tool called "Diagram Helper" I created back in 2014 to design flow charts, but found it can be easily modified to work as a form designer. I also added a small WinForms library to SB and wrote a pre-compiler to lower the object syntax to the normal syntax that the SB compiler understands. Using this simple trick, I allowed using control names as objects to make the code shorter and easier.
So, I ended up with a Visual Small Basic which was too close to Visual Basic or in fact a small version of it, hence I decided to name it Small Visual Basic. 
 
![image](https://github.com/VBAndCs/sVB-Small-Visual-Basic/assets/48354902/8e3975ee-a5b3-4a4f-8cce-eb013fa2f5d6)

![image](https://github.com/VBAndCs/sVB-Small-Visual-Basic/assets/48354902/5dff43e5-ba32-4125-92bc-3f97ccf0aea4)

I published the a sVB prerelease on GitHub in January 2021, and invited the SB team to use the idea to evolve SB, but they didn't seem interested!
For about 5 months, I had no intension to do anything further, but in June 2021, my beloved grandmother Om Wahba died, God rest her soul, so, I decided to complete sVB to honor her. 
Also in that time, my nephew Mohammad Mohsen started to see my YouTube videos about SB, and I wanted to spare him the SB issues I explained above, so I worked hard to release the first sVB stable version in July 2021, which had only 3 controls in the toolbox! 
I kept updating the language over the next two years, until it hit v2.8 with the release of this book.
That was the story of the three years journey from Small Basic to Small Visual Basic.

# It is also a Small Visual Basic .NET!
You can say that:
Small Visual Basic = Small Basic + Visual Basic:
but this in not the whole truth!
sVB is meant to be a middle step in between SB and VB.NET, so, its syntax combines the best of SB, VB, and VB.NET languages!
It is also as easy as Windows Forms, but also has some advanced WPF graphical features!
Let's look a little deeper into that:

# sVB and Small Basic:
sVB is built on top Small Basic IDE, compiler and library.

![image](https://github.com/VBAndCs/sVB-Small-Visual-Basic/assets/48354902/830ea85b-adbc-4f12-b1c3-c4e9b3367937)

In most cases, SB code works in sVB, but there are a few breaking changes that we will explain in the next paragraphs.
sVB also introduces some enhancements in to SB, such as:
1.	SB uses the EndFor and EndWhile keywords. They are still valid in sVB but it allows (and recommends) you to use the Next and Wend keywords instead, to be more consistent with VB.
2.	sVB added many methods to the SB types like Array, Text and File, and made changes to a few of the old methods like making the file methods return True or False instead of "SUCCESS" or "FAILED". You can learn about this in the sVB Library part of this book.
3.	sVB added the TextWindow.WriteLines method to allow you to directly output an array of items each at a line, which becomes very handy when used with the sVB array initializer:
TextWindow.WriteLines({
   "Line1",
   "Line2",
   "Line3"
})
4.	sVB adds many enhancements to the Graphics Window. For example:
a.	In SB you can add only a Button and a TextBox on the graphics window, but in sVB you can add all the controls you see in the sVB toolbox by the using the Controls.AddX methods like Controls.AddCheckBox, Controls.AddComboBox, Controls.AddDatePicker, … etc. These methods return the keys of the added controls, where sVB can deal with them as objects of these controls types, so it is easy to access their properties and methods and add handlers for their various events (See examples provided with each method).
b.	You can create a composite shape using the geometric path and draw it on the graphics window by calling the Shapes.AddGeometricPath method.
c.	Since the graphics window is in fact a normal form like any other form you add to your sVB project, you can use the GraphicsWindow.AsForm method to get the form object that represents it, so you can access more of its properties and methods.

# sVB and Visual Basic:
sVB has many enhancements over SB to make writing apps fast and easy with little code. It brings back the joy and excitement of using vb6 to write RAD applications.
To do that, sVB contains a small WinForms library and a form designer that allows you to add many forms to the project and only one global module.

![image](https://github.com/VBAndCs/sVB-Small-Visual-Basic/assets/48354902/28d69312-c3a6-4b38-86e5-ecdf2ddb3aee)
 
You can drag controls from the toolbox, drop them on the form and double-click any control to get its default event handler created for you, where the ControlName_EventName naming convention is used to recognize the event handlers. 

![image](https://github.com/VBAndCs/sVB-Small-Visual-Basic/assets/48354902/7cfb7492-e7b0-4ee4-bac4-4952f3c3e570)

You can also add event handlers from the upper dropdown lists, by choosing the control name from the left list (say "Button1"), and click the event name from the right list (say "OnClick"), so this sub will be added for you in the code editor:
```VB
Sub Button1_OnClick()
   
EndSub
```

Like SB, sVB is a dynamic language, that uses the same syntax rules of Visual Basic syntax when you use these two statements at the top of the VB code file:
```VB
Option Strict Off 
Option Explicit Off
```

sVB also has some other VB syntax features that are not supported by SB, like:
1.	sVB uses control names as objects to access their properties and methods directly using the dot, like:
TextBox1.Text = "Hello sVB!"
TextBox1 is still a dynamic variable that contains the string key of the control (which is "Form1.TextBox1"), and you can read it, or change it to any value at any time (which of course is not recommended), but the editor auto-completion and intellisense make you feel it is an object, and the sVB compiler translates this object-like syntax to the normal SB syntax like:
```VB
TextBox.SetText(TextBox1, "Hello sVB!")
```

2.	sVB allows you to use Me to refer the current form.
3.	sVB allows you to use the line continuity symbol _ to split a long command line at any position you want (except around the dot). For example:
```VB
TextBox1.Text = "Hello sVB! " + _
    "This is my first program."
```

4.	In SB, all variables are global, which means that the whole file has only one variable scope. sVB acts like VB, where variables defined in a subroutine are local to it, while variables defined at the file level (outside subroutines and functions) are global to the whole file.
5.	sVB allows you to declare Functions that return values.
6.	sVB allows subroutines and functions to declare parameters.
7.	sVB has a ForEach statement to loop through arrays. It is similar to VB For Each.
8.	sVB uses ExitLoop to exit For and While loops, similar to Exit For and Exit While in VB.

But note that there are a few syntax differences between SB/sVB and VB, such as:
1.	sVB uses array[i] to index the array instead of array(i) in VB.
2.	sVB array is actually a dictionary, so you can use string keys instead of numeric indexes, and this is why array indexes don't need to be continuous nor ordered, and can even be negative such as array[-1], because the compiler actually treats that as array["-1"].
3.	The two keyword statements in VB (like For Each, End If, End While and End Sub) are combined into one word in sVB: (ForEach, EndIf, EndWhile and EndSub).

# sVB and Visual Basic .NET:
sVB in fact can be called sVB.NET, because it supports some VB.NET syntax features like:
1.	sVB allows implicit line continuity before or after some symbols like comma, parentheses and arithmetic operators:
```VB
TextBox1.Text = "Hello sVB! " 
    + "This is my first program."
```

2.	sVB uses array initializers {} to add many elements to the array:
```VB
a = {1, 2, 3, 4}
```

3.	sVB supports date and time literals like:
#1/31/2023 13:00:00#
Note that the above literal uses the date format of the English culture.
Date literals are more flexile in sVB than in VB.NET, since sVB supports any valid English culture date format that can be parsed with the .NET Framework DateTime.Parse method, like #31 Jan 2023#. 
sVB even has a time span literal like #+3:15:00# which is not supported by VB.NET!
4.	sVB uses some famous .NET predefined enums like Colors and Keys with auto-completion support:
```VB
me.BackColor = Colors.AliceBlue
```

5.	sVB uses ContinueLoop in For and While loops similar to Continue For and Continue While in VB.NET.
6.	sVB uses the dictionary lockup operator ! to access array keys which is similar to dictionaries in VB.NET:
```VB
student["ID"] = 1
TextBox1.Text = student!ID
```

Also, this operator can be used to add dynamic properties to arrays, like the Expando Object in VB.NET, but sVB provides auto-completion dynamic properties, and you can even use comments to provide intellisense info about them:

![image](https://github.com/VBAndCs/sVB-Small-Visual-Basic/assets/48354902/a7723e19-2821-495f-860c-e863115527fe)
 
7.	sVB uses a VB.NET-like type inference, to detect the variable type from its initial value. This doesn't change the fact that the variable is still a dynamic object, and can hold any value of any type at any time, but inferring the initial variable type allows the compiler to treat it as an object, so you can access its properties and method directly like you do in VB.NET. For example:
```VB
n = 2
x = " Test "
TextBox1.AppendLine(n.Power(3))
TextBox1.AppendLine(x.Trim())
```

The above code is the short form of this code:
```VB
n = 2
x = " Test "
TextBox1.AppendLine(Math.Power(n, 3))
TextBox1.AppendLine(Text.Trim(x))
```

The above example may not be so important to you, but it can be when sVB infers the variable type from an expression or a method return value, especially when the method returns a form or a control, so that sVB can treat the variable as a control, which makes it easier to deal with its methods and properties directly:
```VB
lst = Me.AddListBox("lstTest", 10, 10, 200, 300)
lst.AddItem({1, 2, 3, 4})
```

This somehow puts sVB in between of dynamically-typed and statically-typed languages!
8.	The source code of sVB itself is written in VB.NET, and sVB projects are compiled to IL code, the same as VB.NET and C# projects! This means you can use any reflector to decompile any exe created by sVB back to a C# project!
9.	The sVB WinForms controls and their methods are so close to those of the VB.NET WinForms.
10.	sVB has a small UnitTest framework that allows you to write and run test functions to test your project.
11	sVB can create its own class libraries, in addition to the ability of using VB.NET and C# to create libraries for sVB.

# sVB IDE and VS.NET IDE:
sVB editor is built on top of SB editor, with many enhancements that make it more like a small VS.NET editor, such as:
1.	Code blocks auto completion, like adding Then and EndIf after writing If.
2.	Identifiers highlighting and navigation using Ctrl+Shift+Up or Ctrl+Shift+Down.
3.	Intellisense info about focused words with a link to go to definition.
 
![image](https://github.com/VBAndCs/sVB-Small-Visual-Basic/assets/48354902/41d36c1b-98f3-4aac-97de-00107cd54343)

# sVB and WPF:
sVB is built using the Window Presentation Foundation (WPF) framework, to have advanced IDE features. For example, the form designer allows you to easily create advanced visual effects, like rotating and skewing controls, and the color dialog allows you to use solid brushes, linear and radial gradient brushes, and tile and image brushes to draw the control's boarder, background and foreground, which allows you to change the form and controls shapes. 
For example, when you run the FormShape project from the samples folder, you will see this elliptical form, where its transparent ages really don't exist, so you so can click any other windows through them!
 
![image](https://github.com/VBAndCs/sVB-Small-Visual-Basic/assets/48354902/de8b271a-626c-413a-b74b-53039e49160d)

when you run the ListBox2 project from the samples folder, you will see these elliptical buttons:
 
![image](https://github.com/VBAndCs/sVB-Small-Visual-Basic/assets/48354902/b8981db2-75a4-4ace-b2a0-60f6ec5cc2db)

Besides you can use the Control.SetResourceDictionary and Control.SetStyle methods to load a resource dictionary from an external xaml file, to apply advanced styles on the form and any other controls you want to target!
 
![image](https://github.com/VBAndCs/sVB-Small-Visual-Basic/assets/48354902/eaa9015a-5da5-4dc8-8784-8c2884fcac44)

# sVB for kids:
It is fair to wonder about how to introduce sVB to kids, while it seems very big compared to SB.
In fact sVB can be too easy for kids. All you need is to prepare some attractive images, use the color dialog to display them in labels, and use the animation methods to write some short and funny applications that interests small kids (see the animation and Jerry projects in the samples folder).
 
![image](https://github.com/VBAndCs/sVB-Small-Visual-Basic/assets/48354902/125bc726-7d62-402c-b9cc-0e7dfee9688b)

You can start with codeless apps, use the form designer to display some shapes and images, use the color dialog to change their back colors. In the next step you can add some buttons to the app, each with a single code line that can animate the shape location, color or transparency.
In the next step you can introduce variables, but avoid If statements and loops until the kid gets comfortable with the language.
In fact I published [a course to teach sVB for kids on YouTube](https://www.youtube.com/@smallvisualbasic), but it is in Arabic. It contains about 60 videos (10 minutes in average), and I think it is easy for developers and teachers to view them even they don't understand Arabic, as the visual design and the little sVB code used, with the aid of this book, will be easy to comprehend, so, you can get the idea and build upon. 
I encourage youtubers to introduce such a course in English and other languages, and they are free to reuse any ideas from my course without worrying about copyrights.

# Conclusions:
1.	Small visual Basic is not really SMALL, as you can use it to create large projects containing tens of forms. The Small part of the name is just to reflect that it is built on Small Basic.
2.	sVB combines the best of SB, VB, VB.NET, VS.NET IDE, WinForms and WPF frameworks.
3.	sVB is an easy lightweight version of the .NET platform, with only 6MB zip file to download, and minimum hardware requirements to run!
4.	sVB propose is to be attractive for kids and beginners to learn programming, and to easily move on later to the .NET platform.
5.	By learning sVB, you are only one step forward to learn VB.NET.
6.	sVB trains you on using the Windows Forms framework (WinForms), and it also opens a small window to make you take a look at some amazing WPF capabilities, so anytime later, you may find yourself interested in learning one of the XAML family technologies like WPF, MAUI or WinUI.
