Imports System.IO
Imports System.Media
Imports System.Threading
Imports System.Windows.Media
Imports Microsoft.SmallVisualBasic.Library
Imports Microsoft.SmallVisualBasic.Library.Internal

Namespace WinForms
    ''' <summary>
    ''' The Shape object allows you to add, move and rotate shapes on the Graphics window.
    ''' </summary>
    <SmallVisualBasicType>
    <HideFromIntellisense>
    Public NotInheritable Class ShapesEx

        ''' <summary>
        ''' Gets or sets the text of the shape, if it is a text object.
        ''' </summary>
        <ExProperty>
        Public Shared Function GetText(shapeName As Primitive) As Primitive
            Return Shapes.GetText(shapeName)
        End Function

        Public Shared Sub SetText(shapeName As Primitive, text As Primitive)
            Shapes.SetText(shapeName, text)
        End Sub

        ''' <summary>
        ''' Removes the shape from the Graphics Window.
        ''' </summary>
        <ExMethod>
        Public Shared Sub Remove(shapeName As Primitive)
            Shapes.Remove(shapeName)
        End Sub

        ''' <summary>
        ''' Moves the shape to a new position.
        ''' </summary>
        ''' <param name="x">The x co-ordinate of the new position.</param>
        ''' <param name="y">The y co-ordinate of the new position.</param>
        <ExMethod>
        Public Shared Sub Move(shapeName As Primitive, x As Primitive, y As Primitive)
            Shapes.Move(shapeName, x, y)
        End Sub

        ''' <summary>
        ''' Rotates the shape to the specified angle.
        ''' </summary>
        ''' <param name="angle">The angle to rotate the shape.</param>
        <ExMethod>
        Public Shared Sub Rotate(shapeName As Primitive, angle As Primitive)
            Shapes.Rotate(shapeName, angle)
        End Sub

        ''' <summary>
        ''' Scales the shape using the specified zoom levels. Minimum is 0.1 and maximum is 20.
        ''' </summary>
        ''' <param name="scaleX">The x-axis zoom level.</param>
        ''' <param name="scaleY">The y-axis zoom level.</param>
        <ExMethod>
        Public Shared Sub Zoom(shapeName As Primitive, scaleX As Primitive, scaleY As Primitive)
            Shapes.Zoom(shapeName, scaleX, scaleY)
        End Sub

        ''' <summary>
        ''' Animates a shape to a new position.
        ''' </summary>
        ''' <param name="x">The x co-ordinate of the new position.</param>
        ''' <param name="y">The y co-ordinate of the new position.</param>
        ''' <param name="duration">The time for the animation, in milliseconds.</param>
        <ExMethod>
        Public Shared Sub Animate(shapeName As Primitive, x As Primitive, y As Primitive, duration As Primitive)
            Shapes.Animate(shapeName, x, y, duration)
        End Sub

        ''' <summary>
        ''' Gets or sets the left co-ordinate of the shape.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Double)>
        <ExProperty>
        Public Shared Function GetLeft(shapeName As Primitive) As Primitive
            Return Shapes.GetLeft(shapeName)
        End Function

        Public Shared Sub SetLeft(shapeName As Primitive, value As Primitive)
            Shapes.SetLeft(shapeName, value)
        End Sub

        ''' <summary>
        ''' Gets or sets the top co-ordinate of the shape.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.Double)>
        <ExProperty>
        Public Shared Function GetTop(shapeName As Primitive) As Primitive
            Return Shapes.GetTop(shapeName)
        End Function

        Public Shared Sub SetTop(shapeName As Primitive, value As Primitive)
            Shapes.SetTop(shapeName, value)
        End Sub

        ''' <summary>
        ''' Gets or sets the opacity of a shape. Valid values are between 0 and 100, where 0 is completely transparent and 100 is completely opaque.
        ''' </summary>        
        <WinForms.ReturnValueType(VariableType.Double)>
        <ExProperty>
        Public Shared Function GetOpacity(shapeName As Primitive) As Primitive
            Return Shapes.GetOpacity(shapeName)
        End Function

        Sub SetOpacity(shapeName As Primitive, level As Primitive)
            Shapes.SetOpacity(shapeName, level)
        End Sub

        ''' <summary>
        ''' Hides the shape.
        ''' </summary>
        <ExMethod>
        Public Shared Sub Hide(shapeName As Primitive)
            Shapes.HideShape(shapeName)
        End Sub

        ''' <summary>
        ''' Shows the shape if it is hidden.
        ''' </summary>
        <ExMethod>
        Public Shared Sub Show(shapeName As Primitive)
            Shapes.ShowShape(shapeName)
        End Sub

        ''' <summary>
        ''' Animates the shape on the geometic path, last created using the GeometricPath type.
        ''' </summary>
        ''' <param name="duration">The time for the animation, in milliseconds.</param>
        <ExMethod>
        Private Shared Sub AnimateOnPath(
                   shapeName As Primitive,
                   duration As Primitive)
            Shapes.AnimateOnPath(shapeName, duration)
        End Sub
    End Class
End Namespace
