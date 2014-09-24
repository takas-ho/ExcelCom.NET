Namespace Core
    Public Class Font : Inherits AbstractExcelSubObject : Implements IExcelObject

        Public Sub New(ByVal parent As IExcelObject, ByVal comObject As Object)
            MyBase.New(parent, comObject)
        End Sub

        Public Property Name() As String
            Get
                Return InvokeGetProperty(Of String)("Name")
            End Get
            Set(ByVal value As String)
                InvokeSetProperty("Name", value)
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

    End Class
End Namespace