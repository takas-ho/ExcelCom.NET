Namespace Core

    Public Class Workbook : Inherits AbstractExcelSubObject : Implements IExcelObject

        Public Sub New(ByVal parent As IExcelObject, ByVal comObject As Object)
            MyBase.New(parent, comObject)
        End Sub

        Public Sub Activate()
            InvokeMethod("Activate")
        End Sub

        Private _sheets As Sheets
        Public Function Sheets() As Sheets
            If _sheets Is Nothing Then
                _sheets = New Sheets(Me, InvokeGetProperty("Sheets"))
            End If
            Return _sheets
        End Function

        Public Sub SaveAs(ByVal fileName As String)
            InvokeMethod("SaveAs", fileName)
        End Sub

    End Class
End Namespace