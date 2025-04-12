# Sahla-En 1.0

**Sahla** can simplify coding for children. Currently, it includes Arabic and English dialects, but it is easy to create more dialects. This enables kids to program in their native tongue starting at the age of 6, following a gradual learning path to professionalism:
- **Native Sahla** →  
- **English Sahla** →  
- **Small Basic** →  
- **Small Visual Basic** →  
- **Visual Basic .NET** →  
- **C#**
---

# Create Your Own Native Sahla in a Few Minutes!

It is easy to translate Sahla syntax into your native language, allowing your kids to learn programming without the foreign language barrier. This approach introduces basic programming concepts to little kids and shows them how to apply math to solve small problems.
While you might worry that this could isolate them from the international programming community and mainstream programming languages, Sahla is designed to avoid that:

- **Code Conversion:**  
  You can slso watch [this video](https://youtu.be/VY6VG8nwFhc)
  Sahla allows you to convert native code to the Sahla English syntax. This means kids from all nations can communicate and share their code. It also enables a smooth transition from programming with their native language to programming in English—all using the same IDE and familiar concepts. To convert native code to English, use the **File → Save As** menu and, in the save file window, open the "save as type" list and choose **English Sahla Code Files ("*.shlen")**.

- **Stepping-up to Advanced Languages:**  
  Sahla is very small, and kids will soon feel the need to move on to a "bigger" language. The next best step is **Small Visual Basic (sVB)**, the language used to write Sahla. sVB uses its code library to write, read, play sounds, and perform calculations. It is an easy and powerful educational programming language that comes with over 160 samples and a full PDF reference book, along with a kid programmer book series available on Amazon. Designed as an easy doorway to VB .NET and C#, Sahla also allows you to convert its code to sVB. To do this, use the **File → Save As** menu and, in the save file window, open the "save as type" list and choose **Small Basic Files ("*.sb")**.

---

## Getting Started: Create Your Native Sahla

If you find this interesting, follow these steps to create your native Sahla:

1. **Install sVB:**  
   The source code of Sahla is included in the sVB samples folder. Start by installing sVB on your Windows PC from:  
   [sVB Installer](https://marketplace.visualstudio.com/items?itemName=ModernVBNET.sVBInstaller)

2. **Open the Sample:**  
   After installing sVB, double-click the sVB Samples shortcut on your desktop. Navigate to the `sVB Samples\Sahla Programming Language\Sahla-En` folder and double-click the `Global.sb` file to open it in sVB.

3. **Edit Global File:**  
   At the top of the global file, you will find 64 lines defining the error messages and keywords of Sahla. They are in English, but you can translate them into your own language. **Note:** Do not use any spaces or symbols in keyword names.

4. **Map Letters for Variables:**  
   Use the upper right list to navigate to the `MapLetters` subroutine. This subroutine converts variable names written in your language’s letters to English ones, making the converted code in English Sahla readable. The default code maps Arabic letters to English. Any letter not mapped will be used as-is, so if you want to retain your variable names, you may ignore this subroutine.

5. **Form Design Settings:**  
   Switch to the "Form design" tab and double-click the `frmMain` item in the project files list to open its code file. At the top of this file, set the value of the `RightToLeft` variable to `True` if your language is written from right to left.

6. **Customize Number Formatting:**  
   The `Program.UseLocalCulture` property affects how numbers appear.  
   - Setting it to `false` displays numbers in English.  
   - Setting it to `True` displays numbers in your native language.  
   - Deleting or commenting out this property shows numbers according to the context.  
   
   Experiment with these alternatives to determine the best option for you.

7. **Configure File Extension:**  
   Set the value of the `SahlaFilesExtension` variable. For example:
   - Use `"shl"` for the Arabic dialect.
   - Use `"shlen"` for the English dialect.
   - Follow a similar pattern like `"shlfr"` for a French dialect.

8. **Keyboard Input Language:**  
   Scroll down to the `Form_OnShown` subroutine and uncomment the last line to enable Sahla to switch the keyboard input language to the user's PC local language:
   ```vb
   Program.SwitchKeyboardToLocalLanguage()
9.	Translate the Menu:
Switch back to the form design and right-click any empty area of the form. Click the "Menu designer" command to open it. Use the menu designer to translate menu item names into your language.
10.	Localize Other Windows:
You can easily localize other windows, such as the About window, using the form designer.
 
And that's it! Click frmMain to open it and press F5 to run the project. Enjoy programming using your native Sahla!

