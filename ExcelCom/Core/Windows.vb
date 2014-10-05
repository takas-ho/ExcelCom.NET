Namespace Core

    Public Class Windows : Inherits AbstractExcelSubCollection(Of Window) : Implements IExcelObject

        Public Sub New(ByVal parent As IExcelObject, ByVal comObject As Object)
            MyBase.New(parent, comObject)
        End Sub

        Protected Overrides Function DetectIndex(ByVal item As Window) As Integer
            Return item.Index
        End Function

        Protected Overrides Function DetectIndex(ByVal name As String) As Integer
            Throw New ArgumentException("name引数は使用できない", "name")
        End Function

    End Class
End Namespace