Imports Microsoft.SmallVisualBasic.Library

Namespace Evaluator


    ''' <summary>
    ''' Allows you to evaluate mathematical expressions at runtime.
    ''' </summary>
    <SmallVisualBasicType>
    Public Class Evaluator

        Private Shared _expression As Primitive
        Private Shared expr As Expressions.Expression
        Friend Shared Errors As List(Of [Error])


        ''' <summary>
        ''' Gets or sets the mathematical expression that you want to evaluate. This expression is case-insensitive.
        ''' You can use x as the input variable, and you can send its value as a parameter when you call the Evaluate method.
        ''' The expression can contain arithmitic operators like (+ - * / % Mod ^).
        ''' You can also use the methods of the Math type. For example:
        ''' Evaluator.Expression = "(2 + 3 * x) / Sin(GetRadians(x))"
        ''' For simplicity, you can use Radians instead or GetRadians or even use Rad:
        ''' Evaluator.Expression = "(2 + 3 * x) / Sin(Rad(x))"
        ''' For more info about the math expression syntax, see the sVB docs reference book.
        ''' </summary>        
        Public Shared Property Expression As Primitive
            Get
                Return _expression
            End Get

            Set(value As Primitive)
                _expression = value
                If value.IsEmpty Then
                    expr = Nothing
                    Errors?.Clear()
                Else
                    Dim _parser As New MathParser
                    expr = _parser.Parse(value)
                    Errors = _parser.Errors
                End If
            End Set
        End Property

        ''' <summary>
        ''' Calculates the value of the mathematical expression with x substituted with the given value.
        ''' </summary>
        ''' <param name="x">The value to substitute x with.</param>
        ''' <returns>The value of the mathematical expression at the given x.</returns>
        <WinForms.ReturnValueType(VariableType.Double)>
        Public Shared Function Evaluate(x As Primitive) As Primitive
            Try
                Return If(expr Is Nothing, New Primitive(""), expr.Evaluate(x))
            Catch ex As Exception
                Errors.Add(New [Error](0, 0, ex.Message))
            End Try

            Return ""
        End Function

        ''' <summary>
        ''' Returns a string that contains details about last errors that happened while parsing and evaluating the current expression.
        ''' </summary>
        <WinForms.ReturnValueType(VariableType.String)>
        Public Shared ReadOnly Property LastErrors As Primitive
            Get
                If Errors Is Nothing OrElse Errors.Count = 0 Then
                    Return ""
                End If

                Dim sb As New System.Text.StringBuilder()
                For Each e In Errors
                    sb.AppendLine(e.Description)
                Next
                Return sb.ToString()
            End Get
        End Property
    End Class

End Namespace