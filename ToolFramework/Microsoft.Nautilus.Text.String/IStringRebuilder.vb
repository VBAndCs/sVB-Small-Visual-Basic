Imports System.Collections
Imports System.Collections.Generic
Imports System.Globalization
Imports System.IO

Namespace Microsoft.Nautilus.Text.StringRebuilder
    Public Interface IStringRebuilder
        Inherits IComparable(Of IStringRebuilder), IEnumerable(Of Char), IEnumerable, IEquatable(Of IStringRebuilder)
        ReadOnly Property Length As Integer

        ReadOnly Property LineBreakCount As Integer
        Default ReadOnly Property Item(index As Integer) As Char

        Function GetLineNumberFromPosition(position As Integer) As Integer

        Function GetLineFromLineNumber(lineNumber As Integer) As LineSpan

        Overloads Function CompareTo(other As IStringRebuilder, ignoreCase As Boolean, culture As CultureInfo) As Integer

        Function GetText(span1 As Span) As String

        Function ToCharArray(startIndex As Integer, length1 As Integer) As Char()

        Sub CopyTo(sourceIndex As Integer, destination As Char(), destinationIndex As Integer, count As Integer)

        Sub Write(writer As TextWriter, span1 As Span)

        Function Substring(span1 As Span) As IStringRebuilder

        Function Insert(position As Integer, text As String) As IStringRebuilder

        Function Insert(position As Integer, text As IStringRebuilder) As IStringRebuilder

        Function Delete(span1 As Span) As IStringRebuilder

        Function Replace(span1 As Span, text As String) As IStringRebuilder

        Function Replace(span1 As Span, text As IStringRebuilder) As IStringRebuilder
    End Interface
End Namespace
