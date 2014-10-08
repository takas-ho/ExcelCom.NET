Namespace Core
    Public Class Shape : Inherits AbstractExcelSubObject : Implements IExcelObject

        Public Sub New(ByVal parent As IExcelObject, ByVal comObject As Object)
            MyBase.New(parent, comObject)
        End Sub

        Public Property Height() As Single
            Get
                Return InvokeGetProperty(Of Single)("Height")
            End Get
            Set(ByVal value As Single)
                InvokeSetProperty("Height", value)
            End Set
        End Property

        Public Property Left() As Single
            Get
                Return InvokeGetProperty(Of Single)("Left")
            End Get
            Set(ByVal value As Single)
                InvokeSetProperty("Left", value)
            End Set
        End Property

        Public ReadOnly Property Name() As String
            Get
                Return InvokeGetProperty(Of String)("Name")
            End Get
        End Property

        Public Property Top() As Single
            Get
                Return InvokeGetProperty(Of Single)("Top")
            End Get
            Set(ByVal value As Single)
                InvokeSetProperty("Top", value)
            End Set
        End Property

        Public Property Width() As Single
            Get
                Return InvokeGetProperty(Of Single)("Width")
            End Get
            Set(ByVal value As Single)
                InvokeSetProperty("Width", value)
            End Set
        End Property

    End Class
End Namespace