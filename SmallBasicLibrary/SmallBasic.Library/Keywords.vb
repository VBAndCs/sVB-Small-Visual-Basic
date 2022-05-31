Namespace Library
    ''' <summary>
    ''' Keywords object is a place holder for providing documentation for Small Basic Keywords
    ''' </summary>
    Public NotInheritable Class Keywords
        ''' <summary>
        ''' The For statement allows you to execute a set of statements multiple times.
        ''' </summary>
        ''' <example>
        ''' The following example will print out numbers from 1 to 10
        ''' <code>
        ''' For i = 1 To 10 
        '''   TextWindow.WriteLine(i)
        ''' Next
        ''' </code>
        ''' </example>
        Public Shared Sub [For]()
        End Sub

        ''' <summary>
        ''' Check the For statement for information about the EndFor keyword.
        ''' </summary>
        Public Shared Sub EndFor()
        End Sub

        ''' <summary>
        ''' Check the For statement for information about the Next keyword.
        ''' </summary>
        Public Shared Sub [Next]()
        End Sub

        ''' <summary>
        ''' Check the For statement for information about the To keyword.
        ''' </summary>
        Public Shared Sub [To]()
        End Sub

        ''' <summary>
        ''' The Step keyword is used to specify an increment in the For loop.
        ''' </summary>
        ''' <example>
        ''' The following example will print out odd numbers from 1 to 10
        ''' <code>
        ''' For i = 1 to 10 Step 2
        '''   TextWindow.WriteLine(i)
        ''' Next
        ''' </code>
        ''' </example>
        Public Shared Sub [Step]()
        End Sub

        ''' <summary>
        ''' The If statement allows you to make decisions to do different things.
        ''' </summary>
        ''' <example>
        ''' The following example will print out either "Win" or "Lose" depending on the outcome of the flip.
        ''' <code>
        ''' If flip = "Tail" Then
        '''   TextWindow.WriteLine("Win")
        ''' Else
        '''   TextWindow.WriteLine("Lose")
        ''' EndIf
        ''' </code>
        ''' </example>
        Public Shared Sub [If]()
        End Sub

        ''' <summary>
        ''' Check the If statement for information about the Then keyword.
        ''' </summary>
        Public Shared Sub [Then]()
        End Sub

        ''' <summary>
        ''' Check the If statement for information about the Else keyword.
        ''' </summary>
        Public Shared Sub [Else]()
        End Sub

        ''' <summary>
        ''' The ElseIf keyword helps provide an alternate condition while making decisions using the If statement.
        ''' </summary>
        ''' <example>
        ''' In the following example, we will print out the right greeting based on the time of the day.
        ''' <code>
        ''' If Clock.Hour &lt; 12 Then
        '''   TextWindow.WriteLine("Good Morning")
        ''' ElseIf Clock.Hour &lt; 16 Then
        '''   TextWindow.WriteLine("Good Afternoon")
        ''' ElseIf Clock.Hour &lt; 20 Then
        '''   TextWindow.WriteLine("Good Evening")
        ''' EndIf
        ''' </code>
        ''' </example>
        Public Shared Sub [ElseIf]()
        End Sub

        ''' <summary>
        ''' Check the If statement for information about the EndIf keyword.
        ''' </summary>
        Public Shared Sub [EndIf]()
        End Sub

        ''' <summary>
        ''' The Goto statement allows branching to a new location in the program.  
        ''' </summary>
        ''' <example>
        ''' The following program will print consecutive numbers endlessly.
        ''' <code>
        ''' start:
        ''' TextWindow.WriteLine(i)
        ''' i = i + 1
        ''' Goto start
        ''' </code>
        ''' </example>
        Public Shared Sub [Goto]()
        End Sub

        ''' <summary>
        ''' The Sub (Subroutine) statement allows you to do groups of things with a single call.
        ''' </summary>
        ''' <example>
        ''' The following example defines a subroutine that rings the bell and prints "Win".
        ''' <code>
        ''' Sub Win
        '''   Sound.PlayBellRing()
        '''   TextWindow.WriteLine("Win!")
        ''' EndSub
        ''' </code>
        ''' </example>
        Public Shared Sub [Sub]()
        End Sub

        Public Shared Sub [Function]()
        End Sub

        Public Shared Sub [EndFunction]()
        End Sub
        ''' <summary>
        ''' Check the Sub statement for information about the EndSub keyword.
        ''' </summary>
        Public Shared Sub EndSub()
        End Sub

        ''' <summary>
        ''' The While statement allows you to repeat something until you achieve a desired result.
        ''' </summary>
        ''' <example>
        ''' The following code will print a set of random numbers until one that is greater than 100 is encountered.
        ''' <code>
        ''' While i &lt; 100
        '''   i = Math.GetRandomNumber(150)
        '''   TextWindow.WriteLine(i)
        ''' Wend
        ''' </code>
        ''' </example>
        Public Shared Sub [While]()
        End Sub

        ''' <summary>
        ''' Check the While statement for information about the EndWhile keyword.
        ''' </summary>
        Public Shared Sub EndWhile()
        End Sub

        ''' <summary>
        ''' Check the While statement for information about the Wend keyword.
        ''' </summary>
        Public Shared Sub [Wend]()
        End Sub

        ''' <summary>
        ''' Does a logical computation and returns true if both inputs are true.
        ''' </summary>
        Public Shared Sub [And]()
        End Sub

        ''' <summary>
        ''' Does a logical computation and returns true if either one of the inputs is true.
        ''' </summary>
        Public Shared Sub [Or]()
        End Sub
    End Class
End Namespace
