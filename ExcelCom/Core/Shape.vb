Namespace Core
    Public Class Shape : Inherits AbstractExcelSubObject : Implements IExcelObject

        Public Sub New(ByVal parent As IExcelCollection(Of Shape), ByVal comObject As Object)
            MyBase.New(parent, comObject)
        End Sub

        Public ReadOnly Property Name() As String
            Get
                Return InvokeGetProperty(Of String)("Name")
            End Get
        End Property

    End Class
End Namespace