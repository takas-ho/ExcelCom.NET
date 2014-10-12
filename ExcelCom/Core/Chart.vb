Namespace Core
    Public Class Chart : Inherits AbstractExcelSubObject : Implements IExcelObject

        Public Sub New(ByVal parent As IExcelObject, ByVal comObject As Object)
            MyBase.New(parent, comObject)
        End Sub

        Private _seriesCollection As SeriesCollection
        Public Function SeriesCollection() As SeriesCollection
            If _seriesCollection Is Nothing Then
                _seriesCollection = New SeriesCollection(Me, InvokeMethod("SeriesCollection"))
            End If
            Return _seriesCollection
        End Function

    End Class
End Namespace