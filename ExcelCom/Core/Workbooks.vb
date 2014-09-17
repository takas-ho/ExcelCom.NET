Namespace Core

    Public Class Workbooks : Inherits AbstractExcelSubCollection(Of Workbook) : Implements IExcelObject

        Public Sub New(ByVal parent As IExcelObject, ByVal comObject As Object)
            MyBase.New(parent, comObject)
        End Sub

        Public Function Add() As Workbook
            Dim result As Workbook = New Workbook(Me, InvokeMethod("Add"))
            InternalItems.Add(result)
            Return result
        End Function

        Public Function Open(ByVal fileName As String, Optional ByVal updateLinks As Boolean = False) As Workbook
            Dim result As Workbook = New Workbook(Me, InvokeMethod("Open", fileName, updateLinks))
            InternalItems.Add(result)
            Return result
        End Function

    End Class
End Namespace