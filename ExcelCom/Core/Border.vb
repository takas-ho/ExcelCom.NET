Namespace Core
    Public Class Border : Inherits AbstractExcelSubObject : Implements IExcelObject

        Public Sub New(ByVal parent As IExcelCollection(Of Border), ByVal comObject As Object)
            MyBase.New(parent, comObject)
        End Sub

    End Class
End Namespace