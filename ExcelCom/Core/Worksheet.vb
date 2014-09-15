Namespace Core

    Public Class Worksheet : Inherits AbstractExcelSubObject : Implements IExcelObject

        Public Sub New(ByVal parent As IExcelObject, ByVal comObject As Object)
            MyBase.New(parent, comObject)
        End Sub

        Public Function Cells(ByVal row As Integer, ByVal column As Integer) As Range
            Return New Range(Me, InvokeGetProperty("Cells", RuleUtil.ConvIndexDotNET2VBA(row), RuleUtil.ConvIndexDotNET2VBA(column)))
        End Function

        Public Function Columns() As Range
            Return New Range(Me, InvokeGetProperty("Columns"))
        End Function

        Public Function Range(ByVal rangeStr As String) As Range
            Return InternalRange(rangeStr)
        End Function

        Public Function Range(ByVal startRange As String, ByVal endRange As String) As Range
            Return InternalRange(startRange, endRange)
        End Function

        Public Function Range(ByVal startRange As Range, ByVal endRange As Range) As Range
            Return InternalRange(startRange.ComObject, endRange.ComObject)
        End Function

        Private Function InternalRange(ByVal cell1 As Object, Optional ByVal cell2 As Object = Nothing) As Range
            Dim args As Object() = If(cell2 Is Nothing, New Object() {cell1}, New Object() {cell1, cell2})
            Return New Range(Me, InvokeGetProperty("Range", args))
        End Function

        Public Function Rows() As Range
            Return New Range(Me, InvokeGetProperty("Rows"))
        End Function

        Public Sub [Select]()
            InvokeMethod("Select")
        End Sub

    End Class
End Namespace