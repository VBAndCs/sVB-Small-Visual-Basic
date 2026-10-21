 <h1 align="center" style="font-size: 56px; font-weight: 900; color: darkred;">
  Small Visual Basic
</h1>

<p align="center">
  <img src="https://modernvbnet.gallerycdn.vsassets.io/extensions/modernvbnet/svbinstaller/3.6.0.2/1792538099755/sVB.png"/>
</p>


### ✅ Easy, fast, free, fun, and powerful programming language for kids and beginners.  
### ✅ Combines the best of Small Basic, Visual Basic 6, VB.NET, Python, WinForms, and WPF.  
### ✅ Your natural path from beginner coding to professional development.  
### ✅ Start coding immediately with 180+ ready-to-use samples. Try it now!  
### 📥 [Download sVB from VS Marketplace](https://marketplace.visualstudio.com/items?itemName=ModernVBNET.sVBInstaller)

<p align="center">
  <img src="https://modernvbnet.gallerycdn.vsassets.io/extensions/modernvbnet/svbinstaller/3.6.0.2/1792538099755/sVB_3.6__2.png"/>
</p>

---

## 👧 Start with "Sahla"

<p align="center">
  <img src="https://modernvbnet.gallerycdn.vsassets.io/extensions/modernvbnet/svbinstaller/3.6.0.2/1792538099755/image__18.png"/>
</p>

**[Sahla](https://marketplace.visualstudio.com/items?itemName=ModernVBNET.Sahla)** is a programming language written entirely in sVB:  
- Has a shortcut installed with sVB on your desktop.  
- Designed for very young children (6 years old and above) to be an easy, fun doorway to sVB.  
- Uses natural language syntax, with direct commands and no dots, commas, or parentheses.  
- Supports variables, basic arithmetic operations, if statements, and loops.  
- Provides commands for input, output, and sound, with built-in animal sounds for creating fun samples. 


<p align="center">
  <img src="https://modernvbnet.gallerycdn.vsassets.io/extensions/modernvbnet/svbinstaller/3.6.0.2/1792538099755/Sahla.png"/>
</p>

- Includes basic turtle commands to create simple geometric shapes.  
- Comes with 24 sample programs.
- Supports Arabic, English, and Serbian dialects.  
- Can be easily localized by translating its keywords within the source code provided in the sVB samples folder.
- As easy as Scratch, thanks to natural commands with auto‑completion, animal sounds, and turtle graphics, making the start even easier and more engaging.

---

## ♟️ Play Chess Too!

<p align="center">
  <img src="https://modernvbnet.gallerycdn.vsassets.io/extensions/modernvbnet/svbinstaller/3.6.0.2/1792538099755/Chess.png"/>
</p>

- A shortcut for Chess is installed on your desktop.  
- Written entirely in sVB (see source code in the samples folder).  
- Supports playing against an auto player, and also includes a 2-player mode.  
- Helps train kids in logical and strategic thinking, which is essential for programming.  
- Teachers can use the 2-player mode to organize chess tournaments for kids.  

---

## ✍️ Let's speak sVB!
Small Visual Basic (sVB) is a dynamic version of BASIC with syntax close to natural English.  
It provides:
- **Auto-completion** of keywords,  method and variable names.  
- **Explicit block endings** (`EndIf`, `Next`, `EndSub`) to avoid confusion.
- **Automatic code formatting** for readability.  
- **Automatic type conversion** (no need for casting).  
- Example Code:
```vb
' Ask the user for his name
TextWindow.WriteLine("What's your name?")
Name = TextWindow.ReadText()
TextWindow.Write("Hello " + Name + ". ")

' Ask him for his birth year
TextWindow.WriteLine("What's your birth year?")
BirthYear = TextWindow.ReadNumber()

' Calculate his age
Today = Date.Now
Age = Today.Year - BirthYear
TextWindow.WriteLine("You are " + Age + " years old.")

' Check if he is a kid, a teenager or a grownup
If Age < 14 Then
   TextWindow.WriteLine("You are a kid.")
ElseIf Age < 18 Then
   TextWindow.WriteLine("You are a teenager.")
Else
   TextWindow.WriteLine("You are a grownup.")
EndIf
```

## 🎮 Fun With Turtle Graphics
Draw colorful shapes, polygons, spirals, stars, flowers, and geometric patterns. 

```vb
' Draw a golden circle with the turtle
GW.PenWidth = 10
GW.PenColor = Colors.Gold
Turtle.Size = 32
Turtle.X = 250
Turtle.Y = 300
For I = 1 To 360
   Turtle.Move(2)
   Turtle.Turn(1)
Next
```

<p align="center">
  <img src="https://modernvbnet.gallerycdn.vsassets.io/extensions/modernvbnet/svbinstaller/3.0/1714546030343/image__11.png"/>
</p>

---

## 🎨 Create Games and Visual Applications with the Graphics Window
Move sprites, handle keyboard/mouse events, create animations, and build arcade games.  
Examples include Chess, animations, and drawing programs.  

<p align="center">
  <img src="https://modernvbnet.gallerycdn.vsassets.io/extensions/modernvbnet/svbinstaller/3.0/1714546030343/image__12.png"/>
</p>

---

## 🖥️ Design Real Windows Applications
- Drag-and-Drop Form Designer  
- Create professional applications visually with buttons, menus, text boxes, timers, date pickers, and lists.
<p align="center">
  <img src="https://modernvbnet.gallerycdn.vsassets.io/extensions/modernvbnet/svbinstaller/3.0/1714546030343/image__1.png"/>
</p>

- Double-click the button ob the designer to handle its click event.
<p align="center">
  <img src="https://modernvbnet.gallerycdn.vsassets.io/extensions/modernvbnet/svbinstaller/3.0/1714546030343/image__2.png"/>
</p> 

- Darg and drop images and rotate them by mouse.
<p align="center">
  <img src="https://modernvbnet.gallerycdn.vsassets.io/extensions/modernvbnet/svbinstaller/3.0/1714546030343/image__10.png"/>
</p>

- Easy control animations on the form to cteate fun samples.

<p align="center">
  <img src="https://modernvbnet.gallerycdn.vsassets.io/extensions/modernvbnet/svbinstaller/3.0/1714546030343/image__9.png"/>
</p>

- Suppoerts easy and advanced XAML styles, like you see in the calculator and listbox sample.
- Can create amazing gmaes like Starship, Cars and Ball games (See the games folder ins sVB samples)
<p align="center">
  <img src="https://modernvbnet.gallerycdn.vsassets.io/extensions/modernvbnet/svbinstaller/3.0/1714546030343/image__13.png"/>
</p>

- Supports multi-form projects and modular applications.  
- Can even create sVB code libraries with their own forms, like the Dialogs library (see its source code in the samples folder).  

---

## 📐 Create Powerful Educational Applications

- Solve quadratic equations, decompose quadratic functions, calculate *e* and *π*, and many more (see the Math Samples folder).
- List right triangles with integral side lengths and draw them (see "Right-Angled Triangles" in the Drawing Samples).
- Use turtle graphics to introduce geometry basics to students (see the Turtle Samples folder and refer to the sVB Kid Programmer book series for more examples).
- Use the turtle and graphics window to show the magic of sine and cosine functions (see the Turtle and Drawing folders in sVB samples).
- Solve interesting puzzles (see the Puzzles folder).
- Use the sVB Evaluator to plot math functions (see "Draw Functions" sample in the Drawing folder).


<p align="center">
  <img src="https://modernvbnet.gallerycdn.vsassets.io/extensions/modernvbnet/svbinstaller/3.0/1714546030343/image__16.png"/>
</p>
<p align="center">
  <img src="https://modernvbnet.gallerycdn.vsassets.io/extensions/modernvbnet/svbinstaller/3.0/1714546030343/image__17.png"/>
</p>


- Drawing Samples contain advanced math drawings such as the Mandelbrot Set, Guilloche Pattern, Sierpiński Triangle, and Archimedean Spiral.
- Use the LitDev library to draw 3D objects (see "3DRubicCube" in the Drawing Samples) and to simulate physics (see "Falling Balls" in the Games Samples).

---

## 📚 Learning Resources
- **180+ Ready-to-Run Samples**: Games, puzzles, graphics, animations, Windows Forms, and more.  
- **Comprehensive PDF Reference Book**: Covers IDE, syntax, forms designer, and full API.  
- **Amazon Book Series**: ["Small Visual Basic Kid Programmer"](https://www.amazon.com/dp/B0DNX2ZR4D).  
- **Full Small Basic Compatibility**: Run Small Basic code with minimal adjustments.  

<p align="center">
  <a href="https://www.amazon.com/dp/B0DNX2ZR4D">
    <img src="https://m.media-amazon.com/images/I/81ey+FtB8KL._CLa%7C2893,2214%7C71XQRjHURJL.jpg,71dj-VHkYOL.jpg%7C0,0,1315,2214+1578,0,1315,2214+657,0,1578,2214_._SY300_.jpg"/>
  </a>
</p>

---

## 🖥️ Key Features
### 1. Interactive Code Debugger:
Step through code, set breakpoints, and inspect live variables.  

<p align="center">
  <img src="https://modernvbnet.gallerycdn.vsassets.io/extensions/modernvbnet/svbinstaller/3.0/1714546030343/image__15.png"/>
</p>

### 2. Rich Animation & Game Engines:
Create 2D/3D games, animations, and graphics. Includes sound, keyboard/mouse events, and **LitDev** library integration.  

### 3. Advanced Language Capabilities:
- Functions with parameters and local variables.  
- Inline array initialization.  
- Localized documentation (`SmallVisualBasicLibrary.xx.xml`).  
- Math plotting with **sVB Evaluator**.  

- Multi-form projects and custom libraries.  
- Unit testing with `UnitTest` library.  
- Multi-threading via `Thread.SubToRun`.  

### 4. Enhanced Classic Environments:
Improved **TextWindow**, **GraphicsWindow**, and **Turtle** with modern performance.  

### 5. Natural Path to Professional Development:
#### Sahla → sVB → VB.NET → C# → .NET  

---

## 🆚 Why Choose sVB Over Python?
| Feature | Small Visual Basic (sVB) | Python |
|---------|---------------------------|--------|
| GUI Design | Built-in Form Designer | Requires Tkinter/PyQt |
| Case Sensitivity | Case-insensitive | Case-sensitive |
| Array Indexing | 1-based | 0-based |
| Block Structure | Explicit (`EndIf`, `Next`) | Indentation-sensitive |
| Type Conversion | Automatic | Requires casting (`int()`, `str()`) |
| Loop Syntax | `For I = 1 To 5 ... Next` | `for i in range(1,6):` |
| Professional Path | VB.NET → C# → Visual Studio | Multiple separate paths |

---

## 🌍 Open Source
The sVB source code is written in **VB.NET (WPF targeting .NET Framework 4.5)** and hosted on [GitHub](https://github.com/VBAndCs/sVB-Small-Visual-Basic).  

---
