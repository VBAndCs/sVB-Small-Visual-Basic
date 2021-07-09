Imports System.Windows
Imports System.Windows.Controls
Imports Microsoft.SmallBasic.Documents

Namespace Microsoft.SmallBasic.Shell
    Public Class Ribbon
        Shared Sub New()
            FrameworkElement.DefaultStyleKeyProperty.OverrideMetadata(GetType(Ribbon), New FrameworkPropertyMetadata(GetType(Ribbon)))
        End Sub


        Public Shared ReadOnly ParameterProperty As DependencyProperty = DependencyProperty.RegisterAttached("Parameter", GetType(Object), GetType(Ribbon))

        Public Shared Function GetParameter(target As DependencyObject) As Object
            Return target.GetValue(ParameterProperty)
        End Function

        Public Shared Sub SetParameter(target As DependencyObject, value As Object)
            target.SetValue(ParameterProperty, value)
        End Sub


        Public Shared ReadOnly LargeIconProperty As DependencyProperty = DependencyProperty.RegisterAttached("LargeIcon", GetType(UIElement), GetType(Ribbon))

        Public Shared Function GetLargeIcon(target As DependencyObject) As UIElement
            Return CType(target.GetValue(LargeIconProperty), UIElement)
        End Function

        Public Shared Sub SetLargeIcon(target As DependencyObject, value As UIElement)
            target.SetValue(LargeIconProperty, value)
        End Sub


        Public Shared ReadOnly SmallIconProperty As DependencyProperty =
                        DependencyProperty.RegisterAttached("SmallIcon",
                        GetType(UIElement), GetType(Ribbon))

        Public Shared Function GetSmallIcon(target As DependencyObject) As UIElement
            Return CType(target.GetValue(SmallIconProperty), UIElement)
        End Function

        Public Shared Sub SetSmallIcon(target As DependencyObject, value As UIElement)
            target.SetValue(SmallIconProperty, value)
        End Sub

    End Class
End Namespace
