Namespace Core
    Public Class ChartObjects : Inherits AbstractExcelSubCollection(Of ChartObject) : Implements IExcelObject

        Public Sub New(ByVal parent As IExcelObject, ByVal comObject As Object)
            MyBase.New(parent, comObject)
        End Sub

        Protected Overrides Function DetectIndex(ByVal item As ChartObject) As Integer
            Return item.Index
        End Function

        Protected Overrides Function DetectIndex(ByVal name As String) As Integer
            For i As Integer = 0 To Me.Count - 1
                If name.Equals(InternalItems(i).Name) Then
                    Return i
                End If
            Next
            Return -1
        End Function

        Public Function Add(ByVal Left As Double, ByVal Top As Double, ByVal Width As Double, ByVal Height As Double) As ChartObject
            Dim comObject As Object = InvokeMethod("Add", Left, Top, Width, Height)
            If comObject Is Nothing Then
                Return Nothing
            End If
            Return New ChartObject(Me, comObject)
        End Function

    End Class
End Namespace