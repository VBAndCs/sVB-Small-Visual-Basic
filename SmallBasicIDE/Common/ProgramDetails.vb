Imports System
Imports System.CodeDom.Compiler
Imports System.ComponentModel
Imports System.Diagnostics
Imports System.Xml.Serialization

Namespace Microsoft.SmallVisualBasic.com.smallbasic
    <Serializable>
    <GeneratedCode("System.Xml", "4.0.30319.1")>
    <DebuggerStepThrough>
    <DesignerCategory("code")>
    <XmlType([Namespace]:="http://tempuri.org/")>
    Public Class ProgramDetails
        Private idField As String
        Private titleField As String
        Private descriptionField As String
        Private categoryField As String
        Private myRatingField As Double
        Private ratingField As Double
        Private popularityField As Integer
        Private numberOfRatingsField As Integer

        Public Property Id As String
            Get
                Return idField
            End Get
            Set(value As String)
                idField = value
            End Set
        End Property

        Public Property Title As String
            Get
                Return titleField
            End Get
            Set(value As String)
                titleField = value
            End Set
        End Property

        Public Property Description As String
            Get
                Return descriptionField
            End Get
            Set(value As String)
                descriptionField = value
            End Set
        End Property

        Public Property Category As String
            Get
                Return categoryField
            End Get
            Set(value As String)
                categoryField = value
            End Set
        End Property

        Public Property MyRating As Double
            Get
                Return myRatingField
            End Get
            Set(value As Double)
                myRatingField = value
            End Set
        End Property

        Public Property Rating As Double
            Get
                Return ratingField
            End Get
            Set(value As Double)
                ratingField = value
            End Set
        End Property

        Public Property Popularity As Integer
            Get
                Return popularityField
            End Get
            Set(value As Integer)
                popularityField = value
            End Set
        End Property

        Public Property NumberOfRatings As Integer
            Get
                Return numberOfRatingsField
            End Get
            Set(value As Integer)
                numberOfRatingsField = value
            End Set
        End Property
    End Class
End Namespace
