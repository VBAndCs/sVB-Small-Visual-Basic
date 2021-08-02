Imports System.Drawing
Imports System.Windows.Forms

Namespace Library
    ''' <summary>
    ''' The mouse class provides accessors to get or set the mouse related properties, like the cursor position, pointer, etc.
    ''' </summary>
    <SmallBasicType>
    Public NotInheritable Class Mouse

        ''' <summary>
        ''' Gets or sets the mouse cursor's x co-ordinate.
        ''' </summary>
        Public Shared Property MouseX As Primitive
            Get
                Return New Primitive(System.Windows.Forms.Cursor.Position.X)
            End Get

            Set(Value As Primitive)
                System.Windows.Forms.Cursor.Position = New Point(value, MouseY)
            End Set
        End Property

        ''' <summary>
        ''' Gets or sets the mouse cursor's y co-ordinate.
        ''' </summary>
        Public Shared Property MouseY As Primitive
            Get
                Return New Primitive(System.Windows.Forms.Cursor.Position.Y)
            End Get

            Set(Value As Primitive)
                System.Windows.Forms.Cursor.Position = New Point(MouseX, value)
            End Set
        End Property

        ''' <summary>
        ''' Gets whether or not the left button is pressed.
        ''' </summary>
        Public Shared ReadOnly Property IsLeftButtonDown As Primitive
            Get
                Return New Primitive((Control.MouseButtons And MouseButtons.Left) = MouseButtons.Left)
            End Get
        End Property

        ''' <summary>
        ''' Gets whether or not the right button is pressed.
        ''' </summary>
        Public Shared ReadOnly Property IsRightButtonDown As Primitive
            Get
                Return New Primitive((Control.MouseButtons And MouseButtons.Right) = MouseButtons.Right)
            End Get
        End Property

        ''' <summary>
        ''' Hides the mouse cursor on the screen.
        ''' </summary>
        Public Shared Sub HideCursor()
            GraphicsWindow.SetCursor(System.Windows.Input.Cursors.None)
        End Sub

        ''' <summary>
        ''' Shows the mouse cursors on the screen.
        ''' </summary>
        Public Shared Sub ShowCursor()
            GraphicsWindow.SetCursor(System.Windows.Input.Cursors.Arrow)
        End Sub
    End Class
End Namespace
