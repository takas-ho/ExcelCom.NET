Namespace Core
    Public Class Font : Inherits AbstractExcelSubObject : Implements IExcelObject

#Region "Xl定数"
        Public Enum XlUnderlineStyle
            xlUnderlineStyleDouble = -4119
            xlUnderlineStyleDoubleAccounting = 5
            xlUnderlineStyleNon = -4142
            xlUnderlineStyleSingle = 2
            xlUnderlineStyleSingleAccounting = 4
        End Enum
#End Region
        Public Sub New(ByVal parent As IExcelObject, ByVal comObject As Object)
            MyBase.New(parent, comObject)
        End Sub

        Public Property Bold() As Boolean
            Get
                Return InvokeGetProperty(Of Boolean)("Bold")
            End Get
            Set(ByVal value As Boolean)
                InvokeSetProperty("Bold", value)
            End Set
        End Property

        Public Property Italic() As Boolean
            Get
                Return InvokeGetProperty(Of Boolean)("Italic")
            End Get
            Set(ByVal value As Boolean)
                InvokeSetProperty("Italic", value)
            End Set
        End Property

        Public Property Name() As String
            Get
                Return InvokeGetProperty(Of String)("Name")
            End Get
            Set(ByVal value As String)
                InvokeSetProperty("Name", value)
            End Set
        End Property

        Public Property Shadow() As Boolean
            Get
                Return InvokeGetProperty(Of Boolean)("Shadow")
            End Get
            Set(ByVal value As Boolean)
                InvokeSetProperty("Shadow", value)
            End Set
        End Property

        Public Property Size() As Double
            Get
                Return InvokeGetProperty(Of Double)("Size")
            End Get
            Set(ByVal value As Double)
                InvokeSetProperty("Size", value)
            End Set
        End Property

        Public Property Underline() As XlUnderlineStyle
            Get
                Return InvokeGetProperty(Of XlUnderlineStyle)("Underline")
            End Get
            Set(ByVal value As XlUnderlineStyle)
                InvokeSetProperty("Underline", value)
            End Set
        End Property

    End Class
End Namespace