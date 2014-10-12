Namespace Core
    Public Class Chart : Inherits AbstractExcelSubObject : Implements IExcelObject

        Public Sub New(ByVal parent As IExcelObject, ByVal comObject As Object)
            MyBase.New(parent, comObject)
        End Sub

        Public Function SeriesCollection() As SeriesCollection
            Return New SeriesCollection(Me, InvokeMethod("SeriesCollection"))
        End Function

    End Class
End Namespace