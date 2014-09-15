Namespace Core

    Public Class Workbooks : Inherits AbstractExcelSubObject : Implements IExcelObject

        Public Sub New(ByVal parent As IExcelObject, ByVal comObject As Object)
            MyBase.New(parent, comObject)
        End Sub

        Public Function Add() As Workbook
            Return New Workbook(Me, InvokeMethod("Add"))
        End Function

        Public Function Open(ByVal fileName As String, Optional ByVal updateLinks As Boolean = False) As Workbook
            Return New Workbook(Me, InvokeMethod("Open", fileName, updateLinks))
        End Function

    End Class
End Namespace