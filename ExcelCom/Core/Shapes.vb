Namespace Core
    Public Class Shapes : Inherits AbstractExcelSubCollection(Of Shape) : Implements IExcelObject

        Public Sub New(ByVal parent As IExcelObject, ByVal comObject As Object)
            MyBase.New(parent, comObject)
        End Sub

        Protected Overrides Function DetectIndex(ByVal item As Shape) As Integer
            For i As Integer = 0 To Me.Count - 1
                If item.Name.Equals(InternalItems(i).Name) Then
                    Return i
                End If
            Next
            Return -1
        End Function

        Protected Overrides Function DetectIndex(ByVal name As String) As Integer
            For i As Integer = 0 To Count - 1
                If name.Equals(InternalItems(i).Name) Then
                    Return i
                End If
            Next
            Return -1
        End Function

        Public Function AddLine(ByVal BeginX As Single, ByVal BeginY As Single, ByVal EndX As Single, ByVal EndY As Single) As Shape
            Dim result As Shape = New Shape(Me, InvokeMethod("AddLine", BeginX, BeginY, EndX, EndY))
            InternalItems.Add(result)
            Return result
        End Function

    End Class
End Namespace